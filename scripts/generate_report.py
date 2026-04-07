"""
Agent Pipeline Weekly Report Generator
Reads from Google Sheets, summarizes with Claude, sends via Gmail SMTP.
Email is formatted as a dark-themed 4-column flowchart:
  Trigger / Schedule → Data Sources → AI Agents → Outputs
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
SNAPSHOT_FILE      = "data/agent_snapshot.json"

# ── Google Sheets ───────────────────────────────────────────────────────────────
def fetch_sheet_data() -> list[dict]:
    """Download sheet as CSV (sheet must be publicly viewable)."""
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid=0"
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
            agents.append(agent)
    return agents

# ── Snapshot (change detection) ────────────────────────────────────────────────
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
    """
    Returns
    -------
    new_rows      : agents not in previous snapshot
    stage_changes : agents whose Stage of Completion changed (gain 'previous_stage' key)
    """
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

# ── Claude Summary ──────────────────────────────────────────────────────────────
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

Write a 3–5 sentence summary that:
1. Gives a quick overall status of the pipeline
2. Calls out newly added agents (if any)
3. Calls out any stage changes — especially agents that moved to "In Progress" or "Done"
4. Uses an upbeat, professional tone suitable for an internal team email

No bullet points. Plain paragraph prose only."""

    msg = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=400,
        messages=[{"role": "user", "content": prompt}],
    )
    return msg.content[0].text.strip()

# ══════════════════════════════════════════════════════════════════════════════
#  FLOWCHART EMAIL BUILDER
#  Dark theme, 4-column table layout (email-client compatible)
# ══════════════════════════════════════════════════════════════════════════════

# Node brand colours
NODE_COLORS: dict[str, str] = {
    "meta":       "#1877F2",
    "tiktok":     "#010101",
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
}
DEFAULT_NODE_BG = "#374151"

# Stage card styles  (text-color, card-bg, border-color, label)
STAGE_CARD: dict[str, tuple[str, str, str, str]] = {
    "Done":        ("#3fb950", "#0d2119", "#238636", "Done"),
    "In Progress": ("#e3b341", "#1c1500", "#d29922", "In Progress"),
    "Planned":     ("#58a6ff", "#0f1729", "#388bfd", "Planned"),
}
DEFAULT_CARD = ("#c9d1d9", "#161b22", "#30363d", "Unknown")


def _node_bg(name: str) -> str:
    nl = name.lower()
    for key, color in NODE_COLORS.items():
        if key in nl:
            return color
    return DEFAULT_NODE_BG


def parse_connections(raw: str) -> tuple[list[str], list[str]]:
    """
    'Meta +TikTok +Google Ads -> Claude -> Gmail'
    Returns (sources, outputs) — Claude node is stripped from the middle.
    sources = everything before the first ->
    outputs = everything after the last ->
    """
    if not raw:
        return [], []
    parts = [p.strip() for p in raw.split("->")]
    sources = [n.strip() for n in parts[0].split("+") if n.strip()]
    outputs: list[str] = []
    if len(parts) > 1:
        last = parts[-1]
        # skip if last segment is just "Claude"
        if last.lower() != "claude":
            outputs = [n.strip() for n in last.split("+") if n.strip()]
    return sources, outputs


def _td(content: str, width: str = "", valign: str = "top", extra: str = "") -> str:
    w = f'width="{width}"' if width else ""
    return f'<td {w} valign="{valign}" {extra}>{content}</td>'


def _arrow_td() -> str:
    """Narrow arrow cell between columns."""
    return (
        '<td width="24" valign="middle" align="center" '
        'style="color:#4d5561;font-size:16px;padding:0 2px;">&#8594;</td>'
    )


def _trigger_card(frequency: str) -> str:
    freq = frequency or "Schedule TBD"
    return f"""
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td style="background:#0d1f36;border:1px solid #1f3358;border-radius:7px;
          padding:12px 14px;">
          <div style="font-size:11px;font-weight:700;color:#79c0ff;margin-bottom:4px;">
            &#9881;&#65039; GitHub Actions Cron
          </div>
          <div style="font-size:11px;color:#8b949e;">{freq}</div>
          <div style="font-size:10px;color:#6e7681;margin-top:3px;">
            Manual dispatch override
          </div>
        </td>
      </tr>
    </table>"""


def _source_nodes(sources: list[str]) -> str:
    if not sources:
        return '<div style="font-size:10px;color:#4d5561;font-style:italic;">—</div>'
    rows = ""
    for s in sources:
        bg = _node_bg(s)
        rows += (
            f'<tr><td style="padding-bottom:5px;">'
            f'<table cellpadding="0" cellspacing="0"><tr>'
            f'<td style="background:{bg};width:8px;height:8px;border-radius:50%;'
            f'font-size:1px;">&nbsp;</td>'
            f'<td style="padding-left:7px;font-size:11px;font-weight:600;'
            f'color:#c9d1d9;white-space:nowrap;">{s}</td>'
            f'</tr></table>'
            f'</td></tr>'
        )
    return (
        f'<table width="100%" cellpadding="0" cellspacing="0">'
        f'<tr><td style="background:#161b22;border:1px solid #30363d;'
        f'border-radius:7px;padding:12px 14px;">'
        f'<table cellpadding="0" cellspacing="0">{rows}</table>'
        f'</td></tr></table>'
    )


def _agent_card(agent: dict) -> str:
    color, bg, border, label = STAGE_CARD.get(agent["stage"], DEFAULT_CARD)
    desc = agent["description"]
    desc_html = (
        f'<div style="font-size:11px;color:#8b949e;margin-top:6px;line-height:1.5;">'
        f'{desc[:120] + "…" if len(desc) > 120 else desc}</div>'
        if desc else ""
    )
    return (
        f'<table width="100%" cellpadding="0" cellspacing="0">'
        f'<tr><td style="background:{bg};border:1px solid {border};'
        f'border-radius:7px;padding:13px 15px;">'
        f'<div style="font-size:9px;font-weight:700;letter-spacing:1px;'
        f'text-transform:uppercase;color:{color};margin-bottom:4px;">{label}</div>'
        f'<div style="font-size:12px;font-weight:700;color:#e6edf3;">{agent["name"]}</div>'
        f'{desc_html}'
        f'</td></tr></table>'
    )


def _output_nodes(outputs: list[str]) -> str:
    if not outputs:
        return '<div style="font-size:10px;color:#4d5561;font-style:italic;">—</div>'
    rows = ""
    for o in outputs:
        bg = _node_bg(o)
        rows += (
            f'<tr><td style="padding-bottom:5px;">'
            f'<table cellpadding="0" cellspacing="0" width="100%"><tr>'
            f'<td style="background:{bg};width:8px;height:8px;border-radius:50%;'
            f'font-size:1px;white-space:nowrap;">&nbsp;</td>'
            f'<td style="padding-left:7px;font-size:11px;font-weight:600;'
            f'color:#c9d1d9;">{o}</td>'
            f'</tr></table>'
            f'</td></tr>'
        )
    return (
        f'<table width="100%" cellpadding="0" cellspacing="0">'
        f'<tr><td style="background:#161b22;border:1px solid #30363d;'
        f'border-radius:7px;padding:12px 14px;">'
        f'<table cellpadding="0" cellspacing="0">{rows}</table>'
        f'</td></tr></table>'
    )


def _full_agent_row(agent: dict) -> str:
    sources, outputs = parse_connections(agent["connections"])
    return f"""
    <tr>
      <td width="190" valign="top" style="padding:0 0 18px 0;">
        {_trigger_card(agent["frequency"])}
      </td>
      {_arrow_td()}
      <td width="160" valign="top" style="padding:0 0 18px 0;">
        {_source_nodes(sources)}
      </td>
      {_arrow_td()}
      <td valign="top" style="padding:0 0 18px 0;">
        {_agent_card(agent)}
      </td>
      {_arrow_td()}
      <td width="160" valign="top" style="padding:0 0 18px 0;">
        {_output_nodes(outputs)}
      </td>
    </tr>
    <tr>
      <td colspan="7"
        style="padding:0 0 18px 0;border-bottom:1px solid #21262d;font-size:1px;">
        &nbsp;
      </td>
    </tr>
    <tr><td colspan="7" style="height:18px;"></td></tr>"""


def _compact_agent_row(agent: dict) -> str:
    """Used for In Progress / Planned agents that have no connections data."""
    color, bg, border, label = STAGE_CARD.get(agent["stage"], DEFAULT_CARD)
    sources, outputs = parse_connections(agent["connections"])
    desc = agent["description"]
    desc_html = (
        f'<div style="font-size:11px;color:#8b949e;margin-top:4px;">'
        f'{desc[:100] + "…" if len(desc) > 100 else desc}</div>'
        if desc else (
            '<div style="font-size:10px;color:#4d5561;font-style:italic;margin-top:4px;">'
            'In development — details coming soon</div>'
        )
    )
    trigger_opacity = "0.4" if not agent["frequency"] else "1"
    src_opacity = "0.3" if not sources else "1"
    out_opacity = "0.3" if not outputs else "1"

    return f"""
    <tr>
      <td width="190" valign="top" style="padding:0 0 14px 0;opacity:{trigger_opacity};">
        {_trigger_card(agent["frequency"])}
      </td>
      {_arrow_td()}
      <td width="160" valign="top" style="padding:0 0 14px 0;opacity:{src_opacity};">
        {_source_nodes(sources) if sources else _source_nodes([])}
      </td>
      {_arrow_td()}
      <td valign="top" style="padding:0 0 14px 0;">
        <table width="100%" cellpadding="0" cellspacing="0">
          <tr>
            <td style="background:{bg};border:1px solid {border};
              border-radius:7px;padding:13px 15px;">
              <div style="font-size:9px;font-weight:700;letter-spacing:1px;
                text-transform:uppercase;color:{color};margin-bottom:4px;">{label}</div>
              <div style="font-size:12px;font-weight:700;color:#e6edf3;">{agent["name"]}</div>
              {desc_html}
            </td>
          </tr>
        </table>
      </td>
      {_arrow_td()}
      <td width="160" valign="top" style="padding:0 0 14px 0;opacity:{out_opacity};">
        {_output_nodes(outputs) if outputs else _output_nodes([])}
      </td>
    </tr>"""


def _col_header_row() -> str:
    labels = ["TRIGGER / SCHEDULE", "DATA SOURCES", "AI AGENTS", "OUTPUTS"]
    cols = ""
    for i, lbl in enumerate(labels):
        if i > 0:
            cols += '<td width="24"></td>'  # spacer for arrow col
        cols += (
            f'<td style="font-size:9px;font-weight:700;letter-spacing:2px;'
            f'text-transform:uppercase;color:#6e7681;padding-bottom:14px;">{lbl}</td>'
        )
    return f'<tr>{cols}</tr>'


def _section_header(title: str, color: str = "#6e7681") -> str:
    return (
        f'<tr><td colspan="7" style="padding:8px 0 14px;">'
        f'<div style="font-size:9px;font-weight:700;letter-spacing:2px;'
        f'text-transform:uppercase;color:{color};border-bottom:1px solid #21262d;'
        f'padding-bottom:8px;">{title}</div>'
        f'</td></tr>'
    )


def _this_week_banner(new_rows: list[dict], stage_changes: list[dict]) -> str:
    if not new_rows and not stage_changes:
        return ""

    rows_html = ""
    for a in new_rows:
        color, bg, border, _ = STAGE_CARD.get(a["stage"], DEFAULT_CARD)
        rows_html += (
            f'<tr style="border-bottom:1px solid #1c2128;">'
            f'<td style="padding:8px 12px;font-size:12px;font-weight:600;'
            f'color:#e6edf3;">{a["name"]}</td>'
            f'<td style="padding:8px 12px;">'
            f'<span style="font-size:10px;font-weight:700;padding:2px 8px;'
            f'background:{bg};color:{color};border:1px solid {border};'
            f'border-radius:20px;">NEW</span></td>'
            f'<td style="padding:8px 12px;font-size:11px;color:#8b949e;">'
            f'{a["description"] or "—"}</td>'
            f'</tr>'
        )
    for a in stage_changes:
        pc, pb, pe, _ = STAGE_CARD.get(a["previous_stage"], DEFAULT_CARD)
        nc, nb, ne, _ = STAGE_CARD.get(a["stage"], DEFAULT_CARD)
        rows_html += (
            f'<tr style="border-bottom:1px solid #1c2128;">'
            f'<td style="padding:8px 12px;font-size:12px;font-weight:600;'
            f'color:#e6edf3;">{a["name"]}</td>'
            f'<td style="padding:8px 12px;white-space:nowrap;">'
            f'<span style="font-size:10px;font-weight:700;padding:2px 7px;'
            f'background:{pb};color:{pc};border:1px solid {pe};border-radius:20px;">'
            f'{a["previous_stage"]}</span>'
            f'<span style="color:#6e7681;font-size:12px;margin:0 5px;">&#8594;</span>'
            f'<span style="font-size:10px;font-weight:700;padding:2px 7px;'
            f'background:{nb};color:{nc};border:1px solid {ne};border-radius:20px;">'
            f'{a["stage"]}</span></td>'
            f'<td style="padding:8px 12px;font-size:11px;color:#8b949e;">'
            f'{a["description"] or "—"}</td>'
            f'</tr>'
        )

    count = len(new_rows) + len(stage_changes)
    return f"""
    <table width="100%" cellpadding="0" cellspacing="0"
      style="margin-bottom:28px;background:#0d2119;border:1px solid #238636;
        border-radius:8px;overflow:hidden;">
      <tr>
        <td style="padding:12px 16px;border-bottom:1px solid #1c4a2a;">
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
            <tr style="background:#0a1e12;">
              <th style="padding:8px 12px;text-align:left;font-size:10px;
                font-weight:700;color:#6e7681;letter-spacing:1px;
                text-transform:uppercase;">Agent</th>
              <th style="padding:8px 12px;text-align:left;font-size:10px;
                font-weight:700;color:#6e7681;letter-spacing:1px;
                text-transform:uppercase;">Update</th>
              <th style="padding:8px 12px;text-align:left;font-size:10px;
                font-weight:700;color:#6e7681;letter-spacing:1px;
                text-transform:uppercase;">Description</th>
            </tr>
            {rows_html}
          </table>
        </td>
      </tr>
    </table>"""


def _legend_row() -> str:
    items = [
        ("#0d1f36", "#1f3358", "Trigger / Schedule"),
        ("#161b22", "#30363d", "Data Source / Output"),
        ("#0d2119", "#238636", "Done"),
        ("#1c1500", "#d29922", "In Progress"),
        ("#0f1729", "#388bfd", "Planned"),
    ]
    cells = ""
    for bg, border, label in items:
        cells += (
            f'<td style="padding-right:20px;white-space:nowrap;">'
            f'<table cellpadding="0" cellspacing="0"><tr>'
            f'<td style="background:{bg};border:1px solid {border};width:12px;'
            f'height:12px;border-radius:3px;font-size:1px;">&nbsp;&nbsp;&nbsp;</td>'
            f'<td style="padding-left:7px;font-size:10px;color:#6e7681;">{label}</td>'
            f'</tr></table></td>'
        )
    return (
        f'<table cellpadding="0" cellspacing="0" style="margin-top:32px;">'
        f'<tr>{cells}</tr></table>'
    )


def build_email_html(
    agents: list[dict],
    new_rows: list[dict],
    stage_changes: list[dict],
    summary: str,
    week_str: str,
) -> str:
    done_agents      = [a for a in agents if a["stage"] == "Done"]
    inprog_agents    = [a for a in agents if a["stage"] == "In Progress"]
    planned_agents   = [a for a in agents if a["stage"] == "Planned"]

    done_count    = len(done_agents)
    inprog_count  = len(inprog_agents)
    planned_count = len(planned_agents)
    total_count   = len(agents)

    # ── Flowchart rows ────────────────────────────────────────────────────────
    done_rows = ""
    for a in done_agents:
        done_rows += _full_agent_row(a)

    inprog_rows = ""
    for a in inprog_agents:
        inprog_rows += _compact_agent_row(a)

    planned_rows = ""
    for a in planned_agents:
        planned_rows += _compact_agent_row(a)

    this_week_html = _this_week_banner(new_rows, stage_changes)
    legend_html    = _legend_row()

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Agent Pipeline Report</title>
</head>
<body style="margin:0;padding:0;background:#0d1117;">
<table width="100%" cellpadding="0" cellspacing="0"
  style="background:#0d1117;padding:32px 20px 56px;">
<tr><td align="center">
<table width="760" cellpadding="0" cellspacing="0" style="max-width:760px;width:100%;">

  <!-- ── PAGE HEADER ── -->
  <tr>
    <td style="padding-bottom:28px;text-align:center;">
      <div style="font-size:20px;font-weight:800;color:#e6edf3;
        letter-spacing:1px;text-transform:uppercase;">
        &#129302; R1 Concepts — Agent Pipeline Report
      </div>
      <div style="margin-top:6px;font-size:12px;color:#6e7681;letter-spacing:0.5px;">
        Week of {week_str}
      </div>
    </td>
  </tr>

  <!-- ── STATS BAR ── -->
  <tr>
    <td style="padding-bottom:24px;">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td width="25%" style="padding-right:10px;">
            <table width="100%" cellpadding="0" cellspacing="0">
              <tr><td style="background:#0d2119;border:1px solid #238636;
                border-radius:7px;padding:14px 10px;text-align:center;">
                <div style="font-size:26px;font-weight:800;color:#3fb950;">{done_count}</div>
                <div style="font-size:9px;font-weight:700;color:#3fb950;
                  text-transform:uppercase;letter-spacing:1px;">Done</div>
              </td></tr>
            </table>
          </td>
          <td width="25%" style="padding-right:10px;">
            <table width="100%" cellpadding="0" cellspacing="0">
              <tr><td style="background:#1c1500;border:1px solid #d29922;
                border-radius:7px;padding:14px 10px;text-align:center;">
                <div style="font-size:26px;font-weight:800;color:#e3b341;">{inprog_count}</div>
                <div style="font-size:9px;font-weight:700;color:#e3b341;
                  text-transform:uppercase;letter-spacing:1px;">In&nbsp;Progress</div>
              </td></tr>
            </table>
          </td>
          <td width="25%" style="padding-right:10px;">
            <table width="100%" cellpadding="0" cellspacing="0">
              <tr><td style="background:#0f1729;border:1px solid #388bfd;
                border-radius:7px;padding:14px 10px;text-align:center;">
                <div style="font-size:26px;font-weight:800;color:#58a6ff;">{planned_count}</div>
                <div style="font-size:9px;font-weight:700;color:#58a6ff;
                  text-transform:uppercase;letter-spacing:1px;">Planned</div>
              </td></tr>
            </table>
          </td>
          <td width="25%">
            <table width="100%" cellpadding="0" cellspacing="0">
              <tr><td style="background:#161b22;border:1px solid #30363d;
                border-radius:7px;padding:14px 10px;text-align:center;">
                <div style="font-size:26px;font-weight:800;color:#c9d1d9;">{total_count}</div>
                <div style="font-size:9px;font-weight:700;color:#6e7681;
                  text-transform:uppercase;letter-spacing:1px;">Total</div>
              </td></tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- ── AI SUMMARY ── -->
  <tr>
    <td style="padding-bottom:24px;">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td style="background:#1a1040;border:1px solid #6e40c9;
            border-radius:8px;padding:16px 20px;">
            <div style="font-size:10px;font-weight:700;letter-spacing:1.5px;
              text-transform:uppercase;color:#a371f7;margin-bottom:8px;">
              &#128203; Weekly Summary
            </div>
            <div style="font-size:13px;line-height:1.7;color:#c9d1d9;">
              {summary}
            </div>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- ── THIS WEEK BANNER ── -->
  <tr>
    <td style="padding-bottom:4px;">
      {this_week_html}
    </td>
  </tr>

  <!-- ── FLOWCHART ── -->
  <tr>
    <td>
      <table width="100%" cellpadding="0" cellspacing="0">

        {_col_header_row()}

        <!-- DONE agents (full rows with sources + outputs) -->
        {_section_header("&#9646; Live Agents", "#3fb950") if done_agents else ""}
        {done_rows}

        <!-- IN PROGRESS agents -->
        {_section_header("&#9646; In Progress", "#e3b341") if inprog_agents else ""}
        {inprog_rows}
        {"<tr><td colspan='7' style='height:10px;'></td></tr>" if inprog_agents else ""}

        <!-- PLANNED agents -->
        {_section_header("&#9646; Planned", "#58a6ff") if planned_agents else ""}
        {planned_rows}

      </table>
    </td>
  </tr>

  <!-- ── LEGEND ── -->
  <tr>
    <td>
      {legend_html}
    </td>
  </tr>

  <!-- ── FOOTER ── -->
  <tr>
    <td style="padding-top:32px;text-align:center;font-size:10px;
      color:#30363d;letter-spacing:0.5px;">
      Auto-generated by Agent Tracker &bull; R1 Concepts &bull; Every Monday 9&nbsp;AM
    </td>
  </tr>

</table>
</td></tr>
</table>
</body>
</html>"""


# ── Email sending ───────────────────────────────────────────────────────────────
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


# ── Main ────────────────────────────────────────────────────────────────────────
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
        print(f"   New agents: {[a['name'] for a in new_rows]}")
    if stage_changes:
        print(f"   Stage changes: {[a['name'] + ' (' + a['previous_stage'] + '→' + a['stage'] + ')' for a in stage_changes]}")
    if not new_rows and not stage_changes:
        print("   No changes detected this week")

    print("🤖 Generating Claude summary...")
    summary = generate_summary(agents, new_rows, stage_changes)

    print("🏗️  Building flowchart email...")
    html = build_email_html(agents, new_rows, stage_changes, summary, week_str)

    print("📧 Sending email...")
    send_email(html, week_str)

    print("💾 Saving snapshot...")
    save_snapshot(agents)

    print("✅ Done!")


if __name__ == "__main__":
    main()
