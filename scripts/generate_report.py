"""
Agent Pipeline Weekly Report Generator
Reads from Google Sheets, summarizes with Claude, sends via Gmail SMTP.
Email uses a dark card-based design grouped by stage.
"""

import os
import json
import csv
import io
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
SNAPSHOT_FILE      = "data/agent_snapshot.json"

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
            "name":        row.get("Name", "").strip(),
            "stage":       row.get("Stage of Completion", "").strip(),
            "frequency":   row.get("Frequency", "").strip(),
            "connections": row.get("Connections", "").strip(),
            "description": row.get("Description", "").strip(),
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

# ── Snapshot ───────────────────────────────────────────────────────────────────
def load_snapshot() -> list[dict]:
    if os.path.exists(SNAPSHOT_FILE):
        with open(SNAPSHOT_FILE) as f:
            return json.load(f)
    return []

def save_snapshot(agents: list[dict]) -> None:
    os.makedirs(os.path.dirname(SNAPSHOT_FILE), exist_ok=True)
    with open(SNAPSHOT_FILE, "w") as f:
        json.dump(agents, f, indent=2)

def detect_changes(
    current: list[dict], previous: list[dict]
) -> tuple[list[dict], list[dict]]:
    prev_map = {a["name"]: a["stage"] for a in previous}
    new_rows: list[dict] = []
    stage_changes: list[dict] = []
    for agent in current:
        name = agent["name"]
        if name not in prev_map:
            new_rows.append(agent)
        elif agent["stage"] != prev_map[name]:
            stage_changes.append({**agent, "previous_stage": prev_map[name]})
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
            f"- {a['name']}: moved from {a['previous_stage']} → {a['stage']}"
            for a in stage_changes
        )
        if stage_changes else "None"
    )
    prompt = f"""You are writing a concise weekly update email about AI automation agents for a marketing team.

Current agent pipeline:
{agent_lines}

New agents added this week:
{new_lines}

Agents whose stage changed this week:
{change_lines}

Write a 3-5 sentence summary that:
1. Gives a quick overall status of the pipeline
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

# ══════════════════════════════════════════════════════════════════════════════
#  EMAIL BUILDER — Dark card design, grouped by stage
# ══════════════════════════════════════════════════════════════════════════════

NODE_COLORS: dict[str, str] = {
    "meta":       "#1877F2",
    "tiktok":     "#111111",
    "google ads": "#4285F4",
    "google":     "#4285F4",
    "hubspot":    "#FF7A59",
    "gmail":      "#EA4335",
    "claude":     "#7C3AED",
    "amazon":     "#FF9900",
    "asana":      "#F06A6A",
    "slack":      "#4A154B",
    "microsoft":  "#00A4EF",
    "teams":      "#6264A7",
    "sheet":      "#34A853",
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
    color, bg, border = STAGE.get(agent["stage"], DEFAULT_STAGE)
    freq = agent["frequency"] or "Schedule TBD"
    desc = agent["description"]
    has_connections = bool(agent["connections"])
    flow_opacity = "1" if has_connections else "0.4"

    return f"""
    <table width="100%" cellpadding="0" cellspacing="0"
      style="background:#0d1117;border:1px solid #21262d;border-radius:12px;
        border-left:3px solid {accent};overflow:hidden;">
      <tr>
        <td style="padding:18px 20px;">

          <!-- Name + description -->
          <div style="font-size:14px;font-weight:700;color:#e6edf3;
            margin-bottom:5px;">{agent["name"]}</div>
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

    # 2-column card grid using a table
    # Pair up agents
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
          {right_td.replace("padding-bottom:12px;", "")}
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

    count = len(new_rows) + len(stage_changes)
    rows_html = ""

    for a in new_rows:
        color, bg, border = STAGE.get(a["stage"], DEFAULT_STAGE)
        rows_html += f"""
        <tr style="border-top:1px solid #122a1a;">
          <td style="padding:8px 16px;font-size:12px;font-weight:600;
            color:#e6edf3;">{a["name"]}</td>
          <td style="padding:8px 16px;white-space:nowrap;">
            <span style="font-size:10px;font-weight:700;padding:2px 8px;
              background:{bg};color:{color};border:1px solid {border};
              border-radius:20px;">NEW</span>
          </td>
          <td style="padding:8px 16px;font-size:11px;color:#6e7681;">
            {a["description"] or "—"}</td>
        </tr>"""

    for a in stage_changes:
        pc, pb, pe = STAGE.get(a["previous_stage"], DEFAULT_STAGE)
        nc, nb, ne = STAGE.get(a["stage"], DEFAULT_STAGE)
        rows_html += f"""
        <tr style="border-top:1px solid #122a1a;">
          <td style="padding:8px 16px;font-size:12px;font-weight:600;
            color:#e6edf3;">{a["name"]}</td>
          <td style="padding:8px 16px;white-space:nowrap;">
            <span style="font-size:10px;font-weight:700;padding:2px 7px;
              background:{pb};color:{pc};border:1px solid {pe};
              border-radius:20px;">{a["previous_stage"]}</span>
            <span style="color:#4d5561;font-size:11px;margin:0 4px;">&#8594;</span>
            <span style="font-size:10px;font-weight:700;padding:2px 7px;
              background:{nb};color:{nc};border:1px solid {ne};
              border-radius:20px;">{a["stage"]}</span>
          </td>
          <td style="padding:8px 16px;font-size:11px;color:#6e7681;">
            {a["description"] or "—"}</td>
        </tr>"""

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
            {count} update{"s" if count != 1 else ""}
          </span>
        </td>
      </tr>
      <tr>
        <td>
          <table width="100%" cellpadding="0" cellspacing="0">
            <tr style="background:#071610;">
              <th style="padding:7px 16px;text-align:left;font-size:9px;font-weight:700;
                color:#6e7681;letter-spacing:1px;text-transform:uppercase;">Agent</th>
              <th style="padding:7px 16px;text-align:left;font-size:9px;font-weight:700;
                color:#6e7681;letter-spacing:1px;text-transform:uppercase;">Update</th>
              <th style="padding:7px 16px;text-align:left;font-size:9px;font-weight:700;
                color:#6e7681;letter-spacing:1px;text-transform:uppercase;">Description</th>
            </tr>
            {rows_html}
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
    completed  = [a for a in agents if a["stage"] == "Completed"]
    in_progress = [a for a in agents if a["stage"] == "In Progress"]
    planned    = [a for a in agents if a["stage"] == "Planned"]

    sections = (
        _section_block("Completed",   completed)
        + _section_block("In Progress", in_progress)
        + _section_block("Planned",     planned)
    )

    this_week = _this_week_block(new_rows, stage_changes)

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
  <title>Agent Pipeline Report</title>
</head>
<body style="margin:0;padding:0;background:#080c12;">
<table width="100%" cellpadding="0" cellspacing="0"
  style="background:#080c12;padding:40px 20px 64px;">
<tr><td align="center">
<table width="900" cellpadding="0" cellspacing="0" style="max-width:900px;width:100%;">

  <!-- ── HEADER ── -->
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
          <td align="right" valign="bottom">
            {stats_html}
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <tr><td style="height:28px;"></td></tr>

  <!-- ── SUMMARY ── -->
  <tr>
    <td style="padding-bottom:16px;">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td style="background:#130d2a;border:1px solid #3b2d6e;
            border-left:3px solid #a371f7;border-radius:12px;padding:18px 22px;">
            <div style="font-size:9px;font-weight:700;letter-spacing:1.5px;
              text-transform:uppercase;color:#a371f7;margin-bottom:8px;">
              &#128203; Weekly Summary
            </div>
            <div style="font-size:13px;line-height:1.75;color:#c9d1d9;">
              {summary}
            </div>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- ── THIS WEEK ── -->
  <tr>
    <td style="padding-bottom:32px;">
      {this_week}
    </td>
  </tr>

  <!-- ── SECTIONS ── -->
  <tr>
    <td>
      {sections}
    </td>
  </tr>

  <!-- ── FOOTER ── -->
  <tr>
    <td style="padding-top:16px;border-top:1px solid #161b22;
      text-align:center;font-size:10px;color:#21262d;letter-spacing:0.5px;">
      Auto-generated by Agent Tracker &bull; R1 Concepts &bull; Every Monday 9&nbsp;AM
    </td>
  </tr>

</table>
</td></tr>
</table>
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


# ── Main ───────────────────────────────────────────────────────────────────────
def main() -> None:
    week_str = datetime.datetime.now().strftime("%B %d, %Y")

    print("📥 Fetching Google Sheet data...")
    agents = fetch_sheet_data()
    print(f"   Found {len(agents)} agents")

    print("📂 Loading previous snapshot...")
    previous = load_snapshot()

    print("🔍 Detecting changes vs last week...")
    new_rows, stage_changes = detect_changes(agents, previous)
    if new_rows:
        print(f"   New: {[a['name'] for a in new_rows]}")
    if stage_changes:
        print(f"   Stage changes: {[a['name']+' ('+a['previous_stage']+'→'+a['stage']+')' for a in stage_changes]}")
    if not new_rows and not stage_changes:
        print("   No changes this week")

    print("🤖 Generating Claude summary...")
    summary = generate_summary(agents, new_rows, stage_changes)

    print("🏗️  Building email...")
    html = build_email_html(agents, new_rows, stage_changes, summary, week_str)

    print("📧 Sending email...")
    send_email(html, week_str)

    print("💾 Saving snapshot...")
    save_snapshot(agents)

    print("✅ Done!")


if __name__ == "__main__":
    main()
