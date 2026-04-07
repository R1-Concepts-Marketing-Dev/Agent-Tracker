"""
Agent Pipeline Weekly Report Generator
Reads from Google Sheets, summarizes with Claude, sends via Gmail SMTP.
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
    """Download sheet as CSV (sheet must be publicly viewable or shared via link)."""
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

# ── Snapshot  ──────────────────────────────────────────────────────────────────
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
    Compare current agents to last week's snapshot.

    Returns
    -------
    new_rows      : agents that didn't exist in the previous snapshot
    stage_changes : agents whose Stage of Completion changed;
                    each dict gains a 'previous_stage' key so the email
                    can show  "Planned → In Progress"
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

# ── Flowchart HTML ──────────────────────────────────────────────────────────────
NODE_COLORS: dict[str, tuple[str, str]] = {
    "claude":      ("#7C3AED", "#FFFFFF"),
    "gmail":       ("#EA4335", "#FFFFFF"),
    "google ads":  ("#4285F4", "#FFFFFF"),
    "meta":        ("#1877F2", "#FFFFFF"),
    "tiktok":      ("#010101", "#FFFFFF"),
    "hubspot":     ("#FF7A59", "#FFFFFF"),
    "amazon":      ("#FF9900", "#111827"),
    "asana":       ("#F06A6A", "#FFFFFF"),
    "slack":       ("#4A154B", "#FFFFFF"),
}
DEFAULT_NODE_COLOR = ("#4B5563", "#FFFFFF")

STAGE_STYLES: dict[str, tuple[str, str]] = {
    "Done":        ("#166534", "#DCFCE7"),
    "In Progress": ("#92400E", "#FEF3C7"),
    "Planned":     ("#1E3A5F", "#DBEAFE"),
}
DEFAULT_STAGE_STYLE = ("#374151", "#F3F4F6")


def node_color(name: str) -> tuple[str, str]:
    for key, colors in NODE_COLORS.items():
        if key in name.lower():
            return colors
    return DEFAULT_NODE_COLOR


def parse_connections(raw: str) -> list[list[str]]:
    """'Meta +TikTok ->Claude ->Gmail'  →  [['Meta','TikTok'], ['Claude'], ['Gmail']]"""
    if not raw:
        return []
    stages = []
    for stage in raw.split("->"):
        nodes = [n.strip() for n in stage.split("+") if n.strip()]
        if nodes:
            stages.append(nodes)
    return stages


def render_node(name: str) -> str:
    bg, fg = node_color(name)
    return (
        f'<span style="display:inline-block;padding:5px 12px;background:{bg};color:{fg};'
        f'border-radius:5px;font-size:12px;font-weight:600;white-space:nowrap;">{name}</span>'
    )


def render_stage_nodes(nodes: list[str]) -> str:
    if len(nodes) == 1:
        return render_node(nodes[0])
    inner = "".join(
        f'<div style="margin:2px 0;">{render_node(n)}</div>' for n in nodes
    )
    return f'<div style="display:inline-block;vertical-align:middle;">{inner}</div>'


ARROW = (
    '<span style="display:inline-block;vertical-align:middle;'
    'color:#9CA3AF;font-size:20px;margin:0 6px;">&#8594;</span>'
)


def build_flowchart_html(agents: list[dict]) -> str:
    rows = []
    for agent in agents:
        stages = parse_connections(agent["connections"])
        if not stages:
            continue

        fg, bg = STAGE_STYLES.get(agent["stage"], DEFAULT_STAGE_STYLE)
        badge = (
            f'<span style="margin-left:8px;padding:2px 9px;background:{bg};color:{fg};'
            f'border-radius:12px;font-size:11px;font-weight:700;">{agent["stage"]}</span>'
        )
        freq = (
            f'<span style="margin-left:10px;color:#6B7280;font-size:12px;">'
            f'&#128337; {agent["frequency"]}</span>'
            if agent["frequency"] else ""
        )
        flow_html = ARROW.join(render_stage_nodes(s) for s in stages)

        rows.append(
            f'<tr><td style="padding:14px 16px;border-bottom:1px solid #E5E7EB;">'
            f'<div style="margin-bottom:10px;">'
            f'<strong style="font-size:13px;color:#111827;">{agent["name"]}</strong>'
            f'{badge}{freq}'
            f'</div>'
            f'<div style="line-height:2;">{flow_html}</div>'
            f'</td></tr>'
        )

    if not rows:
        return (
            '<p style="color:#6B7280;font-style:italic;margin:0;">'
            'No connection data available yet.</p>'
        )

    return (
        '<table style="width:100%;border-collapse:collapse;">'
        + "".join(rows)
        + "</table>"
    )


# ── "This Week" section ────────────────────────────────────────────────────────
def _stage_badge(stage: str, small: bool = False) -> str:
    fg, bg = STAGE_STYLES.get(stage, DEFAULT_STAGE_STYLE)
    px = "2px 7px" if small else "3px 10px"
    return (
        f'<span style="padding:{px};background:{bg};color:{fg};'
        f'border-radius:12px;font-size:11px;font-weight:700;">{stage}</span>'
    )


def build_this_week_section(
    new_rows: list[dict], stage_changes: list[dict]
) -> str:
    if not new_rows and not stage_changes:
        return ""

    total_count = len(new_rows) + len(stage_changes)
    parts: list[str] = []

    # ── New agents sub-table ───────────────────────────────────────────────────
    if new_rows:
        rows_html = "".join(
            f'<tr style="border-bottom:1px solid #D1FAE5;">'
            f'<td style="padding:10px 14px;font-weight:600;color:#111827;">{a["name"]}</td>'
            f'<td style="padding:10px 14px;">{_stage_badge(a["stage"])}</td>'
            f'<td style="padding:10px 14px;color:#374151;font-size:13px;">'
            f'{a["description"] or "—"}</td>'
            f'</tr>'
            for a in new_rows
        )
        parts.append(
            f'<div style="margin-bottom:16px;">'
            f'<div style="font-size:13px;font-weight:700;color:#059669;'
            f'margin-bottom:8px;">&#10133; New Agents Added</div>'
            f'<table width="100%" style="border-collapse:collapse;">'
            f'<tr style="background:#10B981;">'
            f'<th style="padding:9px 14px;text-align:left;color:white;font-size:12px;">Agent</th>'
            f'<th style="padding:9px 14px;text-align:left;color:white;font-size:12px;">Stage</th>'
            f'<th style="padding:9px 14px;text-align:left;color:white;font-size:12px;">Description</th>'
            f'</tr>'
            f'{rows_html}'
            f'</table></div>'
        )

    # ── Stage-change sub-table ─────────────────────────────────────────────────
    if stage_changes:
        rows_html = "".join(
            f'<tr style="border-bottom:1px solid #FEF3C7;">'
            f'<td style="padding:10px 14px;font-weight:600;color:#111827;">{a["name"]}</td>'
            f'<td style="padding:10px 14px;white-space:nowrap;">'
            f'{_stage_badge(a["previous_stage"], small=True)}'
            f'<span style="margin:0 6px;color:#9CA3AF;">&#8594;</span>'
            f'{_stage_badge(a["stage"], small=True)}'
            f'</td>'
            f'<td style="padding:10px 14px;color:#374151;font-size:13px;">'
            f'{a["description"] or "—"}</td>'
            f'</tr>'
            for a in stage_changes
        )
        parts.append(
            f'<div>'
            f'<div style="font-size:13px;font-weight:700;color:#D97706;'
            f'margin-bottom:8px;">&#128260; Stage Progressions</div>'
            f'<table width="100%" style="border-collapse:collapse;">'
            f'<tr style="background:#F59E0B;">'
            f'<th style="padding:9px 14px;text-align:left;color:white;font-size:12px;">Agent</th>'
            f'<th style="padding:9px 14px;text-align:left;color:white;font-size:12px;">Progress</th>'
            f'<th style="padding:9px 14px;text-align:left;color:white;font-size:12px;">Description</th>'
            f'</tr>'
            f'{rows_html}'
            f'</table></div>'
        )

    inner = "\n".join(parts)
    return f"""
    <!-- THIS WEEK -->
    <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:32px;">
      <tr>
        <td style="padding-bottom:10px;border-bottom:2px solid #10B981;">
          <span style="font-size:17px;font-weight:700;color:#111827;">&#10024; This Week &nbsp;
            <span style="font-size:12px;font-weight:600;background:#DCFCE7;color:#166534;
              padding:2px 9px;border-radius:12px;">{total_count} update{"s" if total_count != 1 else ""}</span>
          </span>
        </td>
      </tr>
      <tr>
        <td style="padding-top:14px;">{inner}</td>
      </tr>
    </table>"""


# ── Full Email HTML ─────────────────────────────────────────────────────────────
def build_email_html(
    agents: list[dict],
    new_rows: list[dict],
    stage_changes: list[dict],
    summary: str,
    week_str: str,
) -> str:
    done      = sum(1 for a in agents if a["stage"] == "Done")
    in_prog   = sum(1 for a in agents if a["stage"] == "In Progress")
    planned   = sum(1 for a in agents if a["stage"] == "Planned")
    total     = len(agents)
    flowchart = build_flowchart_html(agents)
    this_week = build_this_week_section(new_rows, stage_changes)

    # All-agents status table rows
    status_rows = ""
    for a in agents:
        fg, bg = STAGE_STYLES.get(a["stage"], DEFAULT_STAGE_STYLE)
        desc = a["description"]
        desc_short = (desc[:85] + "…") if len(desc) > 85 else (desc or "—")
        status_rows += (
            f'<tr style="border-bottom:1px solid #E5E7EB;">'
            f'<td style="padding:10px 12px;font-weight:500;color:#111827;font-size:13px;">'
            f'{a["name"]}</td>'
            f'<td style="padding:10px 12px;">'
            f'<span style="padding:3px 10px;background:{bg};color:{fg};'
            f'border-radius:12px;font-size:11px;font-weight:700;">{a["stage"]}</span></td>'
            f'<td style="padding:10px 12px;color:#6B7280;font-size:12px;">'
            f'{a["frequency"] or "—"}</td>'
            f'<td style="padding:10px 12px;color:#374151;font-size:12px;">{desc_short}</td>'
            f'</tr>'
        )

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Agent Pipeline Report</title>
</head>
<body style="margin:0;padding:0;background:#F3F4F6;
  font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#F3F4F6;padding:24px 0;">
<tr><td align="center">
<table width="680" cellpadding="0" cellspacing="0" style="max-width:680px;width:100%;">

  <!-- HEADER -->
  <tr>
    <td style="background:linear-gradient(135deg,#1E3A5F 0%,#2563EB 100%);
      border-radius:12px 12px 0 0;padding:36px 40px;text-align:center;">
      <div style="font-size:26px;font-weight:800;color:white;letter-spacing:-0.5px;">
        &#129302; Agent Pipeline Report
      </div>
      <div style="margin-top:6px;color:rgba(255,255,255,0.75);font-size:14px;">
        Week of {week_str}
      </div>
    </td>
  </tr>

  <!-- BODY -->
  <tr>
    <td style="background:white;padding:36px 40px;
      border-radius:0 0 12px 12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);">

      <!-- STATS ROW -->
      <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:32px;">
        <tr>
          <td width="25%" style="padding:0 8px 0 0;">
            <table width="100%" style="background:#DCFCE7;border-radius:8px;text-align:center;">
              <tr><td style="padding:16px 8px;">
                <div style="font-size:30px;font-weight:800;color:#166534;">{done}</div>
                <div style="font-size:11px;font-weight:700;color:#166534;
                  text-transform:uppercase;letter-spacing:0.6px;">Done</div>
              </td></tr>
            </table>
          </td>
          <td width="25%" style="padding:0 8px;">
            <table width="100%" style="background:#FEF3C7;border-radius:8px;text-align:center;">
              <tr><td style="padding:16px 8px;">
                <div style="font-size:30px;font-weight:800;color:#92400E;">{in_prog}</div>
                <div style="font-size:11px;font-weight:700;color:#92400E;
                  text-transform:uppercase;letter-spacing:0.6px;">In&nbsp;Progress</div>
              </td></tr>
            </table>
          </td>
          <td width="25%" style="padding:0 8px;">
            <table width="100%" style="background:#DBEAFE;border-radius:8px;text-align:center;">
              <tr><td style="padding:16px 8px;">
                <div style="font-size:30px;font-weight:800;color:#1E3A5F;">{planned}</div>
                <div style="font-size:11px;font-weight:700;color:#1E3A5F;
                  text-transform:uppercase;letter-spacing:0.6px;">Planned</div>
              </td></tr>
            </table>
          </td>
          <td width="25%" style="padding:0 0 0 8px;">
            <table width="100%" style="background:#F3F4F6;border-radius:8px;text-align:center;">
              <tr><td style="padding:16px 8px;">
                <div style="font-size:30px;font-weight:800;color:#374151;">{total}</div>
                <div style="font-size:11px;font-weight:700;color:#374151;
                  text-transform:uppercase;letter-spacing:0.6px;">Total</div>
              </td></tr>
            </table>
          </td>
        </tr>
      </table>

      <!-- AI SUMMARY -->
      <table width="100%" cellpadding="0" cellspacing="0"
        style="margin-bottom:32px;background:#EFF6FF;
          border-radius:8px;border-left:4px solid #2563EB;">
        <tr>
          <td style="padding:20px 24px;">
            <div style="font-size:12px;font-weight:700;color:#1D4ED8;
              text-transform:uppercase;letter-spacing:0.6px;margin-bottom:10px;">
              &#128203; Weekly Summary
            </div>
            <div style="font-size:14px;line-height:1.7;color:#1E3A5F;">
              {summary}
            </div>
          </td>
        </tr>
      </table>

      {this_week}

      <!-- ALL AGENTS TABLE -->
      <div style="margin-bottom:32px;">
        <div style="font-size:17px;font-weight:700;color:#111827;
          padding-bottom:10px;margin-bottom:14px;border-bottom:2px solid #E5E7EB;">
          &#128202; All Agents Status
        </div>
        <table width="100%" cellpadding="0" cellspacing="0">
          <tr style="background:#F9FAFB;">
            <th style="padding:10px 12px;text-align:left;font-size:12px;
              font-weight:700;color:#374151;border-bottom:2px solid #E5E7EB;">Agent</th>
            <th style="padding:10px 12px;text-align:left;font-size:12px;
              font-weight:700;color:#374151;border-bottom:2px solid #E5E7EB;">Status</th>
            <th style="padding:10px 12px;text-align:left;font-size:12px;
              font-weight:700;color:#374151;border-bottom:2px solid #E5E7EB;">Frequency</th>
            <th style="padding:10px 12px;text-align:left;font-size:12px;
              font-weight:700;color:#374151;border-bottom:2px solid #E5E7EB;">Description</th>
          </tr>
          {status_rows}
        </table>
      </div>

      <!-- FLOWCHART -->
      <div style="margin-bottom:32px;">
        <div style="font-size:17px;font-weight:700;color:#111827;
          padding-bottom:10px;margin-bottom:14px;border-bottom:2px solid #E5E7EB;">
          &#128279; Agent Connection Flows
        </div>
        {flowchart}
      </div>

      <!-- FOOTER -->
      <table width="100%" cellpadding="0" cellspacing="0"
        style="border-top:1px solid #E5E7EB;">
        <tr>
          <td style="text-align:center;color:#9CA3AF;font-size:12px;padding-top:16px;">
            Auto-generated by Agent Tracker &bull; R1 Concepts &bull; Every Monday 9&nbsp;AM
          </td>
        </tr>
      </table>

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

    print("🏗️  Building email HTML...")
    html = build_email_html(agents, new_rows, stage_changes, summary, week_str)

    print("📧 Sending email...")
    send_email(html, week_str)

    print("💾 Saving snapshot...")
    save_snapshot(agents)

    print("✅ Done!")


if __name__ == "__main__":
    main()
