"""
Agent Pipeline Weekly Report Generator
Reads from Google Sheets, summarizes with Claude, sends via Gmail SMTP.
Email uses a dark card-based design grouped by stage.

Change detection is date-based:
  - "New this week"      = Date Added column is within the last 7 days
  - "Stage change"       = Stage Updated column is within the last 7 days
  (No snapshot file needed — safe to test repeatedly.)
"""

import os
import csv
import io
import json
import smtplib
import datetime
import urllib.request
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import anthropic

# ── Configuration ──────────────────────────────────────────────────────────────
SHEET_ID           = os.environ["GOOGLE_SHEET_ID"]
ANTHROPIC_API_KEY  = os.environ["ANTHROPIC_API_KEY"]
GMAIL_ADDRESS      = os.environ["GMAIL_ADDRESS"]
GMAIL_APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"]
RECIPIENT_EMAIL    = os.environ.get("RECIPIENT_EMAIL", os.environ["GMAIL_ADDRESS"])
SHEET_GIDS         = os.environ.get("SHEET_GIDS", "0")

# How many days back counts as "this week"
LOOKBACK_DAYS = 7

# ── Google Sheets ───────────────────────────────────────────────────────────────
def _normalize_stage(raw: str) -> str:
    s = raw.strip().lower()
    if s in {"completed", "done", "complete", "finished", "live", "deployed",
             "shipped", "launched", "production"}:
        return "Completed"
    if s in {"in progress", "in-progress", "inprogress", "wip", "active",
             "building", "in development", "in dev", "started", "ongoing"}:
        return "In Progress"
    return "Planned"


def _parse_date(raw: str) -> datetime.date | None:
    """Try common date formats. Returns None if blank or unparseable."""
    raw = raw.strip()
    if not raw:
        return None
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d/%m/%Y",
                "%B %d, %Y", "%b %d, %Y", "%d-%b-%Y"):
        try:
            return datetime.datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
    return None


def _is_recent(date_str: str) -> bool:
    """Return True if the date is within the last LOOKBACK_DAYS days."""
    d = _parse_date(date_str)
    if d is None:
        return False
    cutoff = datetime.date.today() - datetime.timedelta(days=LOOKBACK_DAYS)
    return d >= cutoff


def _fetch_csv(gid: str) -> list[dict]:
    url = (
        f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
        f"/export?format=csv&gid={gid.strip()}"
    )
    with urllib.request.urlopen(url) as resp:
        content = resp.read().decode("utf-8")
    reader = csv.DictReader(io.StringIO(content))
    agents = []
    for row in reader:
        agent = {
            "name":          row.get("Name", "").strip(),
            "stage":         row.get("Stage of Completion", "").strip(),
            "frequency":     row.get("Frequency", "").strip(),
            "connections":   row.get("Connections", "").strip(),
            "description":   row.get("Description", "").strip(),
            "date_added":      row.get("Date Added", "").strip(),
            "stage_updated":   row.get("Date Completed", "").strip(),
        }
        if agent["name"]:
            agent["stage"] = _normalize_stage(agent["stage"])
            agents.append(agent)
    return agents


def fetch_sheet_data() -> list[dict]:
    gids = [g.strip() for g in SHEET_GIDS.split(",") if g.strip()]
    print(f"   Fetching {len(gids)} tab(s) — GIDs: {gids}")
    agents = []
    for gid in gids:
        tab_agents = _fetch_csv(gid)
        print(f"   GID {gid}: {len(tab_agents)} agent(s)")
        agents.extend(tab_agents)
    return agents


# ── Date-based change detection ────────────────────────────────────────────────
def detect_changes(agents: list[dict]) -> tuple[list[dict], list[dict]]:
    """
    new_rows      = agents whose Date Added is within last LOOKBACK_DAYS days
    stage_changes = agents whose Stage Updated is within last LOOKBACK_DAYS days
                    (and Stage Updated != Date Added, i.e. it's a real change)
    """
    new_rows: list[dict] = []
    stage_changes: list[dict] = []

    for agent in agents:
        added   = agent["date_added"]
        updated = agent["stage_updated"]

        if _is_recent(added):
            new_rows.append(agent)
        elif _is_recent(updated):
            # Stage changed this week but agent isn't brand-new
            stage_changes.append(agent)

    return new_rows, stage_changes


# ── Claude Summary ─────────────────────────────────────────────────────────────
def generate_summary(
    agents: list[dict],
    new_rows: list[dict],
    stage_changes: list[dict],
) -> str:
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    agent_lines = "\n".join(
        f"- {a['name']} [{a['stage']}]"
        + (f" | runs {a['frequency']}" if a["frequency"] else "")
        + (f" | {a['description']}" if a["description"] else "")
        for a in agents
    )
    new_lines = (
        "\n".join(f"- {a['name']}: {a['description']}" for a in new_rows)
        if new_rows else "None"
    )
    change_lines = (
        "\n".join(
            f"- {a['name']}: stage updated to {a['stage']}"
            for a in stage_changes
        )
        if stage_changes else "None"
    )
    completed_count   = sum(1 for a in agents if a["stage"] == "Completed")
    in_progress_count = sum(1 for a in agents if a["stage"] == "In Progress")
    planned_count     = sum(1 for a in agents if a["stage"] == "Planned")
    total_count       = len(agents)

    prompt = f"""You are writing a concise weekly update email about AI automation agents for a marketing team.

Pipeline counts (use these exact numbers — do not recount from the list):
- Total agents: {total_count}
- Completed: {completed_count}
- In Progress: {in_progress_count}
- Planned: {planned_count}

Full agent list:
{agent_lines}

New agents added this week:
{new_lines}

Agents whose stage changed this week:
{change_lines}

Write a 3-5 sentence summary that:
1. Gives a quick overall status using the exact pipeline counts above
2. Calls out newly added agents (if any)
3. Calls out any stage changes, especially agents that moved to "In Progress" or "Completed"
4. Uses an upbeat, professional tone suitable for an internal team email

No bullet points. Plain paragraph prose only."""

    msg = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=400,
        messages=[{"role": "user", "content": prompt}],
    )
    return msg.content[0].text.strip()


# ── Claude Workflow Steps ──────────────────────────────────────────────────────
def generate_workflow_steps(agents: list[dict]) -> dict[str, list[dict]]:
    """
    Ask Claude to generate step-by-step workflow descriptions for each agent.
    Returns a dict: { agent_name: [ {platform, action}, ... ] }
    Agents with no connections get an empty list.
    """
    # Only process agents that have connections defined
    agents_with_connections = [a for a in agents if a["connections"].strip()]
    if not agents_with_connections:
        return {}

    agent_blocks = "\n\n".join(
        f'Agent: {a["name"]}\n'
        f'Description: {a["description"] or "No description provided"}\n'
        f'Connections: {a["connections"]}'
        for a in agents_with_connections
    )

    prompt = f"""You are analyzing AI automation agents. For each agent below, generate a step-by-step workflow that explains what the agent actually does at each connection point.

Rules:
- Each step must have a "platform" (the tool/service name, short) and an "action" (what it does at that step, 1 sentence max, plain English)
- Derive the steps logically from the description and connections — make them specific and meaningful
- Keep platform names short (e.g. "HubSpot CRM", "Claude AI", "Gmail", "Google Sheets")
- 2–4 steps per agent maximum
- Return ONLY valid JSON — no explanation, no markdown, no code fences

Return format:
{{
  "Agent Name": [
    {{"platform": "Platform Name", "action": "What it does here"}},
    ...
  ],
  ...
}}

Agents:
{agent_blocks}"""

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    msg = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = msg.content[0].text.strip()

    # Strip markdown code fences if present
    if raw.startswith("```"):
        raw = "\n".join(raw.split("\n")[1:])
    if raw.endswith("```"):
        raw = "\n".join(raw.split("\n")[:-1])

    try:
        return json.loads(raw.strip())
    except json.JSONDecodeError:
        print("⚠️  Could not parse workflow JSON, skipping workflow steps")
        return {}


# ══════════════════════════════════════════════════════════════════════════════
#  EMAIL BUILDER — Dark card design, grouped by stage
# ══════════════════════════════════════════════════════════════════════════════

NODE_COLORS: dict[str, str] = {
    # Google ecosystem
    "google ads":  "#4285F4",
    "google":      "#4285F4",
    "gmail":       "#EA4335",
    "sheet":       "#34A853",
    "drive":       "#34A853",
    # Meta / Social
    "meta":        "#1877F2",
    "facebook":    "#1877F2",
    "instagram":   "#E1306C",
    "tiktok":      "#010101",
    "youtube":     "#FF0000",
    "linkedin":    "#0A66C2",
    "twitter":     "#1DA1F2",
    "x.com":       "#1DA1F2",
    # CRM / Marketing
    "hubspot":     "#FF7A59",
    "salesforce":  "#00A1E0",
    "klaviyo":     "#FFD100",
    "mailchimp":   "#FFE01B",
    # AI
    "claude":      "#7C3AED",
    "anthropic":   "#7C3AED",
    "openai":      "#10A37F",
    "chatgpt":     "#10A37F",
    # Communication
    "slack":       "#4A154B",
    "teams":       "#6264A7",
    "microsoft":   "#00A4EF",
    "twilio":      "#F22F46",
    # E-commerce / Ops
    "amazon":      "#FF9900",
    "shopify":     "#96BF48",
    "stripe":      "#635BFF",
    "zapier":      "#FF4A00",
    "make":        "#6D00CC",
    # Project management
    "asana":       "#F06A6A",
    "notion":      "#FFFFFF",
    "jira":        "#0052CC",
    "airtable":    "#18BFFF",
    # R1 Concepts specific
    "r1":          "#DC2626",
    "dfc":         "#2563EB",
    "website":     "#64748B",
    "api":         "#64748B",
}
DEFAULT_NODE_COLOR = "#4B5563"

# Stage styles: (text, bg, border)
STAGE: dict[str, tuple[str, str, str]] = {
    "Completed":   ("#3fb950", "#0d2119", "#238636"),
    "In Progress": ("#e3b341", "#1c1500", "#d29922"),
    "Planned":     ("#58a6ff", "#0f1729", "#388bfd"),
}
DEFAULT_STAGE = ("#c9d1d9", "#161b22", "#30363d")

# Left-border gradient per stage
CARD_BORDER: dict[str, str] = {
    "Completed":   "#3fb950",
    "In Progress": "#e3b341",
    "Planned":     "#58a6ff",
}

# Section rule colours: (title/count color, line color, count bg, count border)
SECTION_STYLE: dict[str, tuple[str, str, str, str]] = {
    "Completed":   ("#3fb950", "#196130", "#0d2119", "#238636"),
    "In Progress": ("#e3b341", "#5a3e00", "#1c1500", "#d29922"),
    "Planned":     ("#58a6ff", "#1a3a6e", "#0f1729", "#388bfd"),
}


def _node_color(name: str) -> str:
    nl = name.lower()
    for key, color in NODE_COLORS.items():
        if key in nl:
            return color
    return DEFAULT_NODE_COLOR


def _extract_platform_names(agents: list[dict]) -> list[str]:
    """Pull every individual platform name from all agent connection strings."""
    names = set()
    for agent in agents:
        raw = agent.get("connections", "").strip()
        if not raw:
            continue
        # Connections are comma-separated (e.g. "HubSpot, Gmail, Google Sheets")
        for part in raw.split(","):
            part = part.strip()
            if part:
                names.add(part)
    return list(names)


def expand_node_colors_with_claude(agents: list[dict]) -> None:
    """
    Finds platform names in the agents' connections that don't match any
    existing NODE_COLORS keyword, then asks Claude (one batched call) to
    assign official brand hex colors for them. Updates NODE_COLORS in place.
    """
    all_names = _extract_platform_names(agents)
    unknown = [n for n in all_names if _node_color(n) == DEFAULT_NODE_COLOR]

    if not unknown:
        print("   All platform colors already mapped — skipping Claude color expansion")
        return

    print(f"   Asking Claude to assign colors for {len(unknown)} unknown platform(s): {unknown}")

    prompt = f"""For each platform/tool/service name below, return its official brand hex color code.
If you're confident in the brand color, use it. If unsure, pick a visually distinct color that suits the platform's category.

Return ONLY a valid JSON object — no explanation, no markdown, no code fences.
Keys = platform name exactly as given, Values = hex color string (e.g. "#FF5733").

Platforms:
{json.dumps(unknown, indent=2)}"""

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    try:
        msg = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=1000,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = msg.content[0].text.strip()
        if raw.startswith("```"):
            raw = "\n".join(raw.split("\n")[1:])
        if raw.endswith("```"):
            raw = "\n".join(raw.split("\n")[:-1])

        color_map: dict[str, str] = json.loads(raw.strip())
        added = 0
        for name, color in color_map.items():
            key = name.lower()
            # Only add if not already covered (don't overwrite known entries)
            if all(key not in existing for existing in NODE_COLORS):
                NODE_COLORS[key] = color
                added += 1
        print(f"   ✅ Added {added} new color(s) to NODE_COLORS")
    except Exception as e:
        print(f"   ⚠️  Claude color expansion failed ({e}) — falling back to default gray")


def _slugify(name: str) -> str:
    """Convert agent name to a URL-safe slug. e.g. 'Weekly Checkup Bot' -> 'weekly-checkup-bot'"""
    import re
    slug = name.lower().strip()
    slug = re.sub(r"[^\w\s-]", "", slug)   # remove special chars
    slug = re.sub(r"[\s_]+", "-", slug)     # spaces/underscores -> hyphens
    slug = re.sub(r"-+", "-", slug)         # collapse multiple hyphens
    return slug.strip("-")


def _parse_connections(raw: str) -> list[str]:
    """Return a flat ordered list of node names from the connections string."""
    if not raw:
        return []
    nodes: list[str] = []
    for part in raw.split("->"):
        for node in part.split("+"):
            n = node.strip()
            if n:
                nodes.append(n)
    return nodes


def _flow_chips(connections: str) -> str:
    """Render the connection string as inline colored chips with arrows."""
    nodes = _parse_connections(connections)
    if not nodes:
        return (
            '<span style="font-size:11px;color:#4d5561;font-style:italic;">'
            'Not yet configured</span>'
        )

    parts: list[str] = []
    for i, node in enumerate(nodes):
        color = _node_color(node)
        chip = (
            f'<span style="display:inline-block;vertical-align:middle;">'
            f'<table cellpadding="0" cellspacing="0" style="display:inline-table;">'
            f'<tr>'
            f'<td style="background:{color};width:7px;height:7px;border-radius:50%;'
            f'font-size:1px;">&nbsp;</td>'
            f'<td style="padding-left:5px;font-size:11px;font-weight:600;'
            f'color:#c9d1d9;white-space:nowrap;">{node}</td>'
            f'</tr></table></span>'
        )
        if i > 0:
            arrow = (
                '<span style="display:inline-block;vertical-align:middle;'
                'color:#4d5561;font-size:12px;margin:0 5px;">&#8594;</span>'
            )
            parts.append(arrow)
        parts.append(chip)

    return "".join(parts)


def _stage_pill(stage: str) -> str:
    color, bg, border = STAGE.get(stage, DEFAULT_STAGE)
    return (
        f'<span style="font-size:10px;font-weight:700;padding:3px 10px;'
        f'background:{bg};color:{color};border:1px solid {border};'
        f'border-radius:20px;white-space:nowrap;">{stage}</span>'
    )


def _agent_card(agent: dict) -> str:
    """Render a single agent as a dark card with left accent border."""
    accent = CARD_BORDER.get(agent["stage"], "#30363d")
    freq = agent["frequency"] or "Schedule TBD"
    desc = agent["description"]
    has_connections = bool(agent["connections"])
    flow_opacity = "1" if has_connections else "0.4"
    is_new = _is_recent(agent.get("date_added", ""))
    is_stage_updated = _is_recent(agent.get("stage_updated", "")) and not is_new
    just_completed = agent["stage"] == "Completed" and (is_new or is_stage_updated)

    new_badge = ""
    if just_completed:
        new_badge = (
            '<span style="font-size:9px;font-weight:700;padding:2px 8px;'
            'background:#0d1b2e;color:#38bdf8;border:1px solid #0ea5e9;'
            'border-radius:20px;margin-left:8px;vertical-align:middle;">&#10003; Published</span>'
        )

    return f"""
    <table width="100%" cellpadding="0" cellspacing="0"
      style="background:#0d1117;border:1px solid #21262d;border-radius:12px;
        border-left:3px solid {accent};overflow:hidden;">
      <tr>
        <td style="padding:18px 20px;">

          <!-- Name + badge -->
          <div style="font-size:14px;font-weight:700;color:#e6edf3;
            margin-bottom:5px;">{agent["name"]}{new_badge}</div>
          <div style="font-size:11px;color:#8b949e;line-height:1.55;
            margin-bottom:14px;">{desc or "&nbsp;"}</div>

          <!-- Divider -->
          <div style="height:1px;background:#161b22;margin-bottom:14px;"></div>

          <!-- Connection flow -->
          <div style="opacity:{flow_opacity};margin-bottom:14px;line-height:2;">
            {_flow_chips(agent["connections"])}
          </div>

          <!-- Divider -->
          <div style="height:1px;background:#161b22;margin-bottom:12px;"></div>

          <!-- Footer: frequency left, stage right -->
          <table width="100%" cellpadding="0" cellspacing="0">
            <tr>
              <td style="font-size:10px;color:#6e7681;">&#128337; {freq}</td>
              <td align="right">{_stage_pill(agent["stage"])}</td>
            </tr>
          </table>

        </td>
      </tr>
    </table>"""


def _section_block(stage: str, agents: list[dict]) -> str:
    """Render a section header + 2-column card grid for one stage."""
    if not agents:
        return ""

    color, line_color, count_bg, count_border = SECTION_STYLE[stage]
    count = len(agents)

    # Section header: line — STAGE  N  — line
    header = f"""
    <table width="100%" cellpadding="0" cellspacing="0"
      style="margin-bottom:16px;">
      <tr>
        <td style="border-bottom:1px solid {line_color};width:40%;font-size:1px;">&nbsp;</td>
        <td style="padding:0 12px;white-space:nowrap;text-align:center;">
          <span style="font-size:9px;font-weight:700;letter-spacing:2px;
            text-transform:uppercase;color:{color};">{stage.upper()}</span>
          &nbsp;
          <span style="font-size:10px;font-weight:700;padding:1px 8px;
            background:{count_bg};color:{color};border:1px solid {count_border};
            border-radius:20px;">{count}</span>
        </td>
        <td style="border-bottom:1px solid {line_color};width:40%;font-size:1px;">&nbsp;</td>
      </tr>
    </table>"""

    # 2-column card grid
    rows_html = ""
    pairs = [agents[i:i+2] for i in range(0, len(agents), 2)]
    for pair in pairs:
        left  = _agent_card(pair[0])
        right = _agent_card(pair[1]) if len(pair) > 1 else ""
        right_td = (
            f'<td width="50%" valign="top" style="padding-left:6px;">{right}</td>'
            if right else
            '<td width="50%" valign="top" style="padding-left:6px;"></td>'
        )
        rows_html += f"""
        <tr>
          <td width="50%" valign="top" style="padding-right:6px;padding-bottom:12px;">
            {left}
          </td>
          {right_td}
        </tr>"""

    grid = f"""
    <table width="100%" cellpadding="0" cellspacing="0"
      style="margin-bottom:32px;">
      {rows_html}
    </table>"""

    return header + grid


def _this_week_block(new_rows: list[dict], stage_changes: list[dict]) -> str:
    if not new_rows and not stage_changes:
        return ""

    # Stage changes get full detail rows; new agents collapse to a summary line
    stage_rows_html = ""
    for a in stage_changes:
        nc, nb, ne = STAGE.get(a["stage"], DEFAULT_STAGE)
        stage_rows_html += f"""
        <tr style="border-top:1px solid #122a1a;">
          <td style="padding:8px 16px;font-size:12px;font-weight:600;
            color:#e6edf3;">{a["name"]}</td>
          <td style="padding:8px 16px;white-space:nowrap;">
            <span style="font-size:10px;font-weight:700;padding:2px 7px;
              background:{nb};color:{nc};border:1px solid {ne};
              border-radius:20px;">&#8594; {a["stage"]}</span>
          </td>
          <td style="padding:8px 16px;font-size:11px;color:#6e7681;">
            {a["description"] or "—"}</td>
        </tr>"""

    # Compact summary row for new agents
    new_summary_row = ""
    if new_rows:
        new_summary_row = f"""
        <tr style="border-top:1px solid #122a1a;">
          <td colspan="3" style="padding:10px 16px;">
            <span style="font-size:11px;font-weight:700;color:#3fb950;">
              &#43; {len(new_rows)} new agent{"s" if len(new_rows) != 1 else ""} added
            </span>
            <span style="font-size:11px;color:#4d5561;margin-left:8px;">
              {", ".join(a["name"] for a in new_rows[:5])}{"…" if len(new_rows) > 5 else ""}
            </span>
          </td>
        </tr>"""

    total = len(new_rows) + len(stage_changes)

    return f"""
    <table width="100%" cellpadding="0" cellspacing="0"
      style="background:#0a1e12;border:1px solid #196130;border-radius:12px;
        overflow:hidden;margin-bottom:16px;">
      <tr>
        <td style="padding:12px 16px;border-bottom:1px solid #122a1a;">
          <span style="font-size:13px;font-weight:700;color:#3fb950;">
            &#10024; This Week
          </span>
          <span style="margin-left:8px;font-size:10px;font-weight:700;
            background:#14532d;color:#3fb950;padding:2px 9px;border-radius:20px;">
            {total} update{"s" if total != 1 else ""}
          </span>
        </td>
      </tr>
      <tr>
        <td>
          <table width="100%" cellpadding="0" cellspacing="0">
            {new_summary_row}
            {f'''<tr style="background:#071610;">
              <th style="padding:7px 16px;text-align:left;font-size:9px;font-weight:700;
                color:#6e7681;letter-spacing:1px;text-transform:uppercase;">Agent</th>
              <th style="padding:7px 16px;text-align:left;font-size:9px;font-weight:700;
                color:#6e7681;letter-spacing:1px;text-transform:uppercase;">Stage Change</th>
              <th style="padding:7px 16px;text-align:left;font-size:9px;font-weight:700;
                color:#6e7681;letter-spacing:1px;text-transform:uppercase;">Description</th>
            </tr>{stage_rows_html}''' if stage_changes else ""}
          </table>
        </td>
      </tr>
    </table>"""


def build_email_html(
    agents: list[dict],
    new_rows: list[dict],
    stage_changes: list[dict],
    summary: str,
    week_str: str,
) -> str:
    completed   = [a for a in agents if a["stage"] == "Completed"]
    in_progress = [a for a in agents if a["stage"] == "In Progress"]
    planned     = [a for a in agents if a["stage"] == "Planned"]

    def _stat(num, label, top_color, bg_color, text_color):
        return (
            f'<td width="25%" style="padding:0 4px;">'
            f'<div style="background:{bg_color};border-top:3px solid {top_color};'
            f'border-radius:8px 8px 0 0;padding:14px 12px;text-align:center;">'
            f'<div style="font-size:26px;font-weight:800;color:{top_color};">{num}</div>'
            f'<div style="font-size:9px;font-weight:700;color:{text_color};'
            f'text-transform:uppercase;letter-spacing:1px;margin-top:2px;">{label}</div>'
            f'</div></td>'
        )

    stats_html = (
        '<table width="100%" cellpadding="0" cellspacing="0"><tr>'
        + _stat(len(completed),   "Completed",   "#3fb950", "rgba(63,185,80,0.15)",   "rgba(63,185,80,0.7)")
        + _stat(len(in_progress), "In Progress", "#e3b341", "rgba(227,179,65,0.15)",  "rgba(227,179,65,0.7)")
        + _stat(len(planned),     "Planned",     "#58a6ff", "rgba(88,166,255,0.15)",  "rgba(88,166,255,0.7)")
        + _stat(len(agents),      "Total",       "rgba(255,255,255,0.2)", "rgba(255,255,255,0.08)", "rgba(255,255,255,0.3)")
        + '</tr></table>'
    )

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Agent Pipeline Report</title>
</head>
<body style="margin:0;padding:0;background:#f0f2f5;">
<table width="100%" cellpadding="0" cellspacing="0"
  style="background:#f0f2f5;padding:40px 20px 64px;">
<tr><td align="center">
<table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;">

  <!-- ── HERO HEADER ── -->
  <tr>
    <td style="background:#1a1040;border-radius:16px 16px 0 0;padding:36px 36px 0;">

      <!-- Top label -->
      <div style="font-size:10px;font-weight:700;letter-spacing:3px;text-transform:uppercase;
        color:rgba(255,255,255,0.4);margin-bottom:12px;">R1 Concepts &nbsp;&middot;&nbsp; Weekly Update</div>

      <!-- Title row -->
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td>
            <div style="font-size:28px;font-weight:800;color:#fff;line-height:1.1;
              letter-spacing:-0.5px;">Agent Pipeline<br>
              <span style="color:#a78bfa;">Report</span>
            </div>
            <div style="font-size:12px;color:rgba(255,255,255,0.35);margin-top:8px;">
              Week of {week_str}
            </div>
          </td>
          <td align="right" valign="top">
            <div style="font-size:48px;opacity:0.12;line-height:1;">&#129302;</div>
          </td>
        </tr>
      </table>

      <!-- Stats strip -->
      <div style="margin-top:28px;">
        {stats_html}
      </div>
    </td>
  </tr>

  <!-- ── WHITE BODY ── -->
  <tr>
    <td style="background:#ffffff;padding:32px 36px 36px;border-radius:0 0 16px 16px;">

      <!-- Summary label -->
      <div style="font-size:10px;font-weight:700;letter-spacing:2px;text-transform:uppercase;
        color:#a78bfa;margin-bottom:10px;">&#128203; Weekly Summary</div>

      <!-- Summary text -->
      <div style="font-size:14px;line-height:1.8;color:#374151;margin-bottom:28px;
        border-left:3px solid #ede9fe;padding-left:16px;">
        {summary}
      </div>

      <!-- View Full Report Button -->
      <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:32px;">
        <tr>
          <td align="center">
            <a href="https://ben-westreich.github.io/Agent-Tracker/"
              style="display:inline-block;background:#1a1040;
                color:#fff;font-size:13px;font-weight:700;text-decoration:none;
                padding:14px 40px;border-radius:50px;letter-spacing:0.3px;
                border:2px solid #a78bfa;">
              View Full Report &nbsp;&#8594;
            </a>
          </td>
        </tr>
      </table>

      <!-- Footer -->
      <div style="height:1px;background:#f3f4f6;margin-bottom:20px;"></div>
      <div style="text-align:center;font-size:10px;color:#9ca3af;letter-spacing:0.5px;">
        Auto-generated by Agent Tracker &bull; R1 Concepts &bull; Every Friday 12&nbsp;PM PST
      </div>

    </td>
  </tr>

</table>
</td></tr>
</table>
</body>
</html>"""


def _full_report_card(agent: dict, steps: list[dict]) -> str:
    """Render a full report card with numbered workflow steps."""
    accent = CARD_BORDER.get(agent["stage"], "#30363d")
    freq   = agent["frequency"] or "Schedule TBD"
    desc   = agent["description"]

    is_new          = _is_recent(agent.get("date_added", ""))
    is_stage_updated = _is_recent(agent.get("stage_updated", "")) and not is_new
    just_completed  = agent["stage"] == "Completed" and (is_new or is_stage_updated)

    published_badge = ""
    if just_completed:
        published_badge = (
            ' <span style="font-size:9px;font-weight:700;padding:2px 8px;'
            'background:#0d1b2e;color:#38bdf8;border:1px solid #0ea5e9;'
            'border-radius:20px;vertical-align:middle;">&#10003; Published</span>'
        )

    color, bg, border = STAGE.get(agent["stage"], DEFAULT_STAGE)
    stage_pill = (
        f'<span style="font-size:10px;font-weight:700;padding:3px 10px;'
        f'background:{bg};color:{color};border:1px solid {border};'
        f'border-radius:20px;white-space:nowrap;">{agent["stage"]}</span>'
    )

    # Build workflow steps HTML
    if steps:
        steps_html = ""
        for i, step in enumerate(steps):
            platform = step.get("platform", "")
            action   = step.get("action", "")
            dot_color = _node_color(platform)
            is_last  = i == len(steps) - 1
            connector = "" if is_last else (
                '<div style="margin-left:11px;width:1px;height:20px;'
                'background:#21262d;"></div>'
            )
            steps_html += f"""
            <div style="display:flex;gap:12px;align-items:flex-start;">
              <div style="display:flex;flex-direction:column;align-items:center;flex-shrink:0;">
                <div style="width:22px;height:22px;border-radius:50%;
                  background:#1c2128;border:1px solid #30363d;
                  display:flex;align-items:center;justify-content:center;
                  font-size:9px;font-weight:700;color:#6e7681;">{i+1}</div>
                {connector}
              </div>
              <div style="padding-top:3px;{'padding-bottom:16px;' if not is_last else ''}">
                <div style="display:flex;align-items:center;gap:6px;margin-bottom:3px;">
                  <div style="width:7px;height:7px;border-radius:50%;
                    background:{dot_color};flex-shrink:0;"></div>
                  <span style="font-size:11px;font-weight:600;color:#c9d1d9;">{platform}</span>
                </div>
                <div style="font-size:10px;color:#6e7681;line-height:1.5;">{action}</div>
              </div>
            </div>"""

        workflow_section = f"""
        <div style="font-size:9px;font-weight:700;letter-spacing:1.5px;
          text-transform:uppercase;color:#4d5561;margin-bottom:12px;">Workflow</div>
        {steps_html}"""
    else:
        workflow_section = """
        <div style="font-size:9px;font-weight:700;letter-spacing:1.5px;
          text-transform:uppercase;color:#4d5561;margin-bottom:10px;">Workflow</div>
        <div style="background:#0f1117;border:1px dashed #21262d;border-radius:8px;
          padding:12px;text-align:center;">
          <span style="font-size:10px;color:#4d5561;font-style:italic;">
            Not yet configured
          </span>
        </div>"""

    return f"""
    <div style="background:#0d1117;border:1px solid #21262d;border-radius:12px;
      border-left:3px solid {accent};overflow:hidden;margin-bottom:12px;">
      <div style="padding:18px 20px 14px;">

        <!-- Name + badge -->
        <div style="font-size:14px;font-weight:700;color:#e6edf3;margin-bottom:5px;">
          {agent["name"]}{published_badge}
        </div>
        <div style="font-size:11px;color:#8b949e;line-height:1.55;margin-bottom:14px;">
          {desc or "&nbsp;"}
        </div>

        <!-- Divider -->
        <div style="height:1px;background:#161b22;margin-bottom:14px;"></div>

        <!-- Workflow steps -->
        {workflow_section}

        <!-- Divider -->
        <div style="height:1px;background:#161b22;margin:14px 0 12px;"></div>

        <!-- Footer -->
        <div style="display:flex;align-items:center;justify-content:space-between;">
          <span style="font-size:10px;color:#6e7681;">&#128337; {freq}</span>
          {stage_pill}
        </div>

      </div>
    </div>"""


def _full_report_section(stage: str, agents: list[dict], workflow_map: dict) -> str:
    """Render a section with full report cards in a 2-column grid."""
    if not agents:
        return ""

    color, line_color, count_bg, count_border = SECTION_STYLE[stage]
    count = len(agents)

    header = f"""
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:16px;">
      <div style="flex:1;height:1px;background:{line_color};"></div>
      <span style="font-size:9px;font-weight:700;letter-spacing:2px;
        text-transform:uppercase;color:{color};">{stage.upper()}</span>
      <span style="font-size:10px;font-weight:700;padding:1px 8px;
        background:{count_bg};color:{color};border:1px solid {count_border};
        border-radius:20px;">{count}</span>
      <div style="flex:1;height:1px;background:{line_color};"></div>
    </div>"""

    pairs = [agents[i:i+2] for i in range(0, len(agents), 2)]
    grid_rows = ""
    for pair in pairs:
        left_steps  = workflow_map.get(pair[0]["name"], [])
        left_card   = _full_report_card(pair[0], left_steps)
        right_card  = ""
        if len(pair) > 1:
            right_steps = workflow_map.get(pair[1]["name"], [])
            right_card  = _full_report_card(pair[1], right_steps)

        grid_rows += f"""
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;">
          <div>{left_card}</div>
          <div>{right_card}</div>
        </div>"""

    return header + f'<div style="margin-bottom:32px;">{grid_rows}</div>'


def build_full_report_html(
    agents: list[dict],
    new_rows: list[dict],
    stage_changes: list[dict],
    summary: str,
    week_str: str,
    workflow_map: dict = None,
) -> str:
    """Full card-based report for GitHub Pages."""
    if workflow_map is None:
        workflow_map = {}
    completed   = [a for a in agents if a["stage"] == "Completed"]
    in_progress = [a for a in agents if a["stage"] == "In Progress"]
    planned     = [a for a in agents if a["stage"] == "Planned"]

    sections = (
        _full_report_section("Completed",   completed,   workflow_map)
        + _full_report_section("In Progress", in_progress, workflow_map)
        + _full_report_section("Planned",     planned,     workflow_map)
    )

    def _stat(num, label, bg, border, color):
        return (
            f'<td style="padding-right:8px;">'
            f'<table cellpadding="0" cellspacing="0">'
            f'<tr><td style="background:{bg};border:1px solid {border};'
            f'border-radius:8px;padding:10px 16px;text-align:center;min-width:72px;">'
            f'<div style="font-size:22px;font-weight:800;color:{color};">{num}</div>'
            f'<div style="font-size:9px;font-weight:700;color:{color};'
            f'text-transform:uppercase;letter-spacing:1px;margin-top:2px;">{label}</div>'
            f'</td></tr></table></td>'
        )

    stats_html = (
        '<table cellpadding="0" cellspacing="0"><tr>'
        + _stat(len(completed),   "Completed",   "#0d2119", "#238636", "#3fb950")
        + _stat(len(in_progress), "In Progress", "#1c1500", "#d29922", "#e3b341")
        + _stat(len(planned),     "Planned",     "#0f1729", "#388bfd", "#58a6ff")
        + _stat(len(agents),      "Total",       "#161b22", "#30363d", "#e6edf3")
        + '</tr></table>'
    )

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Agent Pipeline Report — {week_str}</title>
</head>
<body style="margin:0;padding:0;background:#080c12;">
<table width="100%" cellpadding="0" cellspacing="0"
  style="background:#080c12;padding:40px 20px 64px;">
<tr><td align="center">
<table width="900" cellpadding="0" cellspacing="0" style="max-width:900px;width:100%;">

  <tr>
    <td style="padding-bottom:32px;border-bottom:1px solid #161b22;">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td valign="bottom">
            <div style="font-size:20px;font-weight:800;color:#e6edf3;
              letter-spacing:-0.5px;">&#129302; Agent Pipeline Report</div>
            <div style="font-size:12px;color:#6e7681;margin-top:4px;">
              R1 Concepts &bull; Week of {week_str}
            </div>
          </td>
          <td align="right" valign="bottom">{stats_html}</td>
        </tr>
      </table>
    </td>
  </tr>

  <tr><td style="height:28px;"></td></tr>

  <tr>
    <td style="padding-bottom:24px;">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td style="background:#130d2a;border:1px solid #3b2d6e;
            border-left:3px solid #a371f7;border-radius:12px;padding:18px 22px;">
            <div style="font-size:9px;font-weight:700;letter-spacing:1.5px;
              text-transform:uppercase;color:#a371f7;margin-bottom:8px;">
              &#128203; Weekly Summary
            </div>
            <div style="font-size:13px;line-height:1.75;color:#c9d1d9;">{summary}</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <tr><td>{sections}</td></tr>

  <tr>
    <td style="padding-top:16px;border-top:1px solid #161b22;
      text-align:center;font-size:10px;color:#21262d;letter-spacing:0.5px;">
      Auto-generated by Agent Tracker &bull; R1 Concepts &bull; Every Friday 12&nbsp;PM PST
    </td>
  </tr>

</table>
</td></tr>
</table>
</body>
</html>"""


# ── Individual Agent Page Builder ─────────────────────────────────────────────
def build_agent_page_html(agent: dict, steps: list[dict], week_str: str) -> str:
    """Build a standalone HTML page for a single agent."""
    slug        = _slugify(agent["name"])
    accent      = CARD_BORDER.get(agent["stage"], "#30363d")
    freq        = agent["frequency"] or "Schedule TBD"
    desc        = agent["description"] or "No description provided."
    connections = _parse_connections(agent["connections"])

    stage       = agent["stage"]
    s_color, s_bg, s_border = STAGE.get(stage, DEFAULT_STAGE)
    stage_dot_map = {"Completed": "#3fb950", "In Progress": "#e3b341", "Planned": "#58a6ff"}
    stage_dot = stage_dot_map.get(stage, "#8b949e")

    date_added     = agent.get("date_added", "") or "—"
    date_completed = agent.get("stage_updated", "") or "—"

    # ── Horizontal workflow track ──
    if steps:
        step_cards = ""
        for i, step in enumerate(steps):
            platform  = step.get("platform", "Unknown")
            action    = step.get("action", "")
            dot_color = _node_color(platform)
            num       = i + 1

            step_cards += f"""
            <div style="display:flex;align-items:center;gap:0;flex-shrink:0;">
              <div style="background:#1c2128;border:1px solid #30363d;border-radius:10px;
                padding:20px 16px;text-align:center;width:140px;position:relative;">
                <!-- Step number badge -->
                <div style="position:absolute;top:-10px;left:50%;transform:translateX(-50%);
                  width:20px;height:20px;border-radius:50%;background:#0d1117;
                  border:1px solid #30363d;display:flex;align-items:center;justify-content:center;
                  font-size:10px;font-weight:700;color:#8b949e;">{num}</div>
                <!-- Colored dot -->
                <div style="width:36px;height:36px;border-radius:50%;
                  background:{dot_color}22;border:2px solid {dot_color}55;
                  display:flex;align-items:center;justify-content:center;
                  margin:0 auto 10px;">
                  <div style="width:12px;height:12px;border-radius:50%;
                    background:{dot_color};"></div>
                </div>
                <div style="font-size:12px;font-weight:700;color:#f0f6fc;
                  margin-bottom:8px;line-height:1.3;">{platform}</div>
                <div style="font-size:10px;color:#6e7681;line-height:1.5;">{action}</div>
              </div>
              {"<!-- arrow --><div style='color:#30363d;padding:0 8px;font-size:22px;'>&#8594;</div>" if i < len(steps) - 1 else ""}
            </div>"""

        workflow_html = f"""
        <div style="overflow-x:auto;padding-bottom:8px;">
          <div style="display:flex;align-items:flex-start;min-width:max-content;gap:0;">
            {step_cards}
          </div>
        </div>"""
    else:
        workflow_html = """
        <div style="background:#1c2128;border:1px dashed #30363d;border-radius:10px;
          padding:24px;text-align:center;">
          <span style="font-size:12px;color:#4d5561;font-style:italic;">
            Workflow not yet configured
          </span>
        </div>"""

    # ── Connections chips ──
    if connections:
        chips = ""
        for i, node in enumerate(connections):
            color = _node_color(node)
            if i > 0:
                chips += '<span style="color:#30363d;font-size:16px;padding:0 4px;">&#8594;</span>'
            chips += (
                f'<span style="display:inline-flex;align-items:center;gap:6px;'
                f'background:#1c2128;border:1px solid #30363d;border-radius:20px;'
                f'padding:5px 12px;font-size:12px;font-weight:500;color:#c9d1d9;">'
                f'<span style="width:8px;height:8px;border-radius:50%;'
                f'background:{color};display:inline-block;"></span>{node}</span>'
            )
        connections_html = f'<div style="display:flex;flex-wrap:wrap;align-items:center;gap:8px;">{chips}</div>'
    else:
        connections_html = '<span style="font-size:12px;color:#4d5561;font-style:italic;">None configured</span>'

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{agent["name"]} — R1 Agent Tracker</title>
  <style>
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{ background: #0d1117; color: #c9d1d9;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      min-height: 100vh; }}
    .topnav {{ background: #161b22; border-bottom: 1px solid #21262d;
      padding: 0 32px; height: 52px; display: flex; align-items: center;
      justify-content: space-between; position: sticky; top: 0; z-index: 100; }}
    .back-btn {{ display: flex; align-items: center; gap: 6px; color: #8b949e;
      text-decoration: none; font-size: 13px; font-weight: 500; transition: color .15s; }}
    .back-btn:hover {{ color: #c9d1d9; }}
    .r1-logo {{ font-size: 12px; font-weight: 700; letter-spacing: 2px;
      color: #DC2626; text-transform: uppercase; }}
    .content {{ max-width: 900px; margin: 0 auto; padding: 36px 32px 64px;
      display: flex; flex-direction: column; gap: 24px; }}
    .card {{ background: #161b22; border: 1px solid #21262d; border-radius: 12px; overflow: hidden; }}
    .card-header {{ padding: 13px 20px; border-bottom: 1px solid #21262d; }}
    .card-label {{ font-size: 10px; font-weight: 700; letter-spacing: 1.5px;
      text-transform: uppercase; color: #8b949e; }}
    .card-body {{ padding: 20px; }}
    .stats-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 14px; }}
    .stat-tile {{ background: #161b22; border: 1px solid #21262d;
      border-radius: 10px; padding: 18px 20px; }}
    .stat-label {{ font-size: 10px; font-weight: 700; letter-spacing: 1.5px;
      text-transform: uppercase; color: #8b949e; margin-bottom: 8px; }}
    .stat-value {{ font-size: 15px; font-weight: 600; color: #f0f6fc; }}
    .footer {{ text-align: center; padding: 24px; font-size: 11px; color: #30363d;
      border-top: 1px solid #161b22; letter-spacing: 0.5px; }}
  </style>
</head>
<body>

<!-- Nav -->
<nav class="topnav">
  <div style="display:flex;align-items:center;gap:20px;">
    <a href="../../index.html" class="back-btn">
      <svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor">
        <path d="M7.78 12.53a.75.75 0 01-1.06 0L2.47 8.28a.75.75 0 010-1.06l4.25-4.25a.75.75 0 011.06 1.06L4.81 7h7.44a.75.75 0 010 1.5H4.81l2.97 2.97a.75.75 0 010 1.06z"/>
      </svg>
      Dashboard
    </a>
    <div style="width:1px;height:18px;background:#30363d;"></div>
    <span style="font-size:13px;color:#8b949e;">
      Agent Tracker &rsaquo; <strong style="color:#c9d1d9;">{agent["name"]}</strong>
    </span>
  </div>
  <div class="r1-logo">R1 Concepts</div>
</nav>

<!-- Hero -->
<div style="background:linear-gradient(160deg,#1a1040 0%,#0d1117 60%);
  border-bottom:1px solid #21262d;border-left:4px solid {accent};">
  <div style="max-width:900px;margin:0 auto;padding:44px 32px 36px;">
    <div style="display:flex;align-items:flex-start;justify-content:space-between;
      gap:16px;margin-bottom:12px;">
      <h1 style="font-size:28px;font-weight:700;color:#f0f6fc;line-height:1.2;">
        {agent["name"]}
      </h1>
      <span style="display:inline-flex;align-items:center;gap:6px;padding:5px 14px;
        border-radius:20px;font-size:12px;font-weight:600;letter-spacing:.5px;
        white-space:nowrap;flex-shrink:0;
        background:{s_bg};border:1px solid {s_border};color:{s_color};">
        <span style="width:7px;height:7px;border-radius:50%;
          background:{stage_dot};display:inline-block;"></span>
        {stage}
      </span>
    </div>
    <div style="display:flex;flex-wrap:wrap;gap:20px;font-size:13px;color:#8b949e;">
      <span>&#128337; {freq}</span>
      <span>&#128197; Added {date_added}</span>
      {"<span>&#10003; Completed " + date_completed + "</span>" if date_completed != "—" else ""}
    </div>
  </div>
</div>

<!-- Content -->
<div class="content">

  <!-- Description -->
  <div class="card">
    <div class="card-header"><div class="card-label">📋 &nbsp;Description</div></div>
    <div class="card-body">
      <p style="font-size:15px;line-height:1.8;color:#c9d1d9;">{desc}</p>
    </div>
  </div>

  <!-- Workflow -->
  <div class="card">
    <div class="card-header"><div class="card-label">⚙️ &nbsp;Workflow</div></div>
    <div class="card-body">
      {workflow_html}
    </div>
  </div>

  <!-- Connections -->
  <div class="card">
    <div class="card-header"><div class="card-label">🔗 &nbsp;Connections</div></div>
    <div class="card-body">
      {connections_html}
    </div>
  </div>

  <!-- Stats -->
  <div class="stats-grid">
    <div class="stat-tile">
      <div class="stat-label">Stage</div>
      <div class="stat-value" style="color:{s_color};">{stage}</div>
    </div>
    <div class="stat-tile">
      <div class="stat-label">Frequency</div>
      <div class="stat-value">{freq}</div>
    </div>
    <div class="stat-tile">
      <div class="stat-label">Platforms</div>
      <div class="stat-value">{len(connections)} connected</div>
    </div>
  </div>

</div>

<div class="footer">
  Auto-generated by Agent Tracker &bull; R1 Concepts &bull; Every Friday 12 PM PST
</div>

</body>
</html>"""


# ── Email sending ──────────────────────────────────────────────────────────────
def send_email(html: str, week_str: str) -> None:
    recipients = [r.strip() for r in RECIPIENT_EMAIL.split(",") if r.strip()]
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"🤖 Agent Pipeline Report — Week of {week_str}"
    msg["From"]    = GMAIL_ADDRESS
    msg["To"]      = ", ".join(recipients)
    msg.attach(MIMEText(html, "html"))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_ADDRESS, recipients, msg.as_string())
    print(f"✅ Email sent to: {', '.join(recipients)}")


# ── Workflow cache (save/load so refresh runs don't need Claude) ───────────────
WORKFLOW_CACHE_PATH = "docs/workflow_cache.json"

def save_workflow_cache(workflow_map: dict) -> None:
    os.makedirs("docs", exist_ok=True)
    with open(WORKFLOW_CACHE_PATH, "w", encoding="utf-8") as f:
        json.dump(workflow_map, f, indent=2)
    print(f"   💾 Workflow cache saved ({len(workflow_map)} agents)")

def load_workflow_cache() -> dict:
    if os.path.exists(WORKFLOW_CACHE_PATH):
        with open(WORKFLOW_CACHE_PATH, encoding="utf-8") as f:
            data = json.load(f)
        print(f"   📂 Loaded workflow cache ({len(data)} agents)")
        return data
    print("   ⚠️  No workflow cache found — pages will show placeholder workflow")
    return {}


# ── Main ───────────────────────────────────────────────────────────────────────
def main() -> None:
    # REFRESH_ONLY=true  → rebuild pages only, no Claude calls, no email
    refresh_only = os.environ.get("REFRESH_ONLY", "false").lower() == "true"
    week_str = datetime.datetime.now().strftime("%B %d, %Y")

    print("📥 Fetching Google Sheet data...")
    agents = fetch_sheet_data()
    print(f"   Found {len(agents)} agents")

    if refresh_only:
        # ── Refresh mode: no Claude, no email ──────────────────────────────
        print("🔄 Refresh-only mode — skipping Claude calls and email")

        print("📂 Loading cached workflow steps...")
        workflow_map = load_workflow_cache()

        print("🏗️  Building full report (GitHub Pages)...")
        new_rows, stage_changes = detect_changes(agents)
        full_html = build_full_report_html(agents, new_rows, stage_changes, "", week_str, workflow_map)

        print("💾 Saving full report to docs/index.html...")
        os.makedirs("docs", exist_ok=True)
        with open("docs/index.html", "w", encoding="utf-8") as f:
            f.write(full_html)

        print("📄 Building individual agent pages...")
        os.makedirs("docs/agents", exist_ok=True)
        for agent in agents:
            slug      = _slugify(agent["name"])
            steps     = workflow_map.get(agent["name"], [])
            page_html = build_agent_page_html(agent, steps, week_str)
            with open(f"docs/agents/{slug}.html", "w", encoding="utf-8") as f:
                f.write(page_html)
        print(f"   ✅ Refreshed {len(agents)} agent page(s)")

    else:
        # ── Full weekly mode: Claude + email ───────────────────────────────
        print("🎨 Expanding platform color map...")
        expand_node_colors_with_claude(agents)

        print(f"🔍 Detecting changes (last {LOOKBACK_DAYS} days)...")
        new_rows, stage_changes = detect_changes(agents)
        if new_rows:
            print(f"   New: {[a['name'] for a in new_rows]}")
        if stage_changes:
            print(f"   Stage updates: {[a['name'] for a in stage_changes]}")
        if not new_rows and not stage_changes:
            print("   No changes this week")

        print("🤖 Generating Claude summary...")
        summary = generate_summary(agents, new_rows, stage_changes)

        print("🤖 Generating workflow steps for full report...")
        workflow_map = generate_workflow_steps(agents)
        print(f"   Generated workflows for {len(workflow_map)} agent(s)")

        print("💾 Saving workflow cache for refresh runs...")
        save_workflow_cache(workflow_map)

        print("🏗️  Building email (simplified)...")
        email_html = build_email_html(agents, new_rows, stage_changes, summary, week_str)

        print("🏗️  Building full report (GitHub Pages)...")
        full_html = build_full_report_html(agents, new_rows, stage_changes, summary, week_str, workflow_map)

        print("💾 Saving full report to docs/index.html...")
        os.makedirs("docs", exist_ok=True)
        with open("docs/index.html", "w", encoding="utf-8") as f:
            f.write(full_html)

        print("📄 Building individual agent pages...")
        os.makedirs("docs/agents", exist_ok=True)
        for agent in agents:
            slug      = _slugify(agent["name"])
            steps     = workflow_map.get(agent["name"], [])
            page_html = build_agent_page_html(agent, steps, week_str)
            with open(f"docs/agents/{slug}.html", "w", encoding="utf-8") as f:
                f.write(page_html)
        print(f"   ✅ Generated {len(agents)} agent page(s) in docs/agents/")

        print("📧 Sending email...")
        send_email(email_html, week_str)

    print("✅ Done!")


if __name__ == "__main__":
    main()
