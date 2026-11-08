"""Render the README images headless with Playwright, from the real run output in examples/<run>/.

    python scripts/render_media.py [local-run]   (needs: pip install playwright && playwright install chromium)

media/run.png        the terminal run
media/excel.png      the Excel table rows it wrote
media/hold-draft.png the calendar hold and the reply left in Drafts
"""
import html
import json
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parent.parent
RUN = ROOT / "examples" / (sys.argv[1] if len(sys.argv) > 1 else "local-run")
MEDIA = ROOT / "media"
PREFIX = sys.argv[2] if len(sys.argv) > 2 else ""
MONO = "'SF Mono', Menlo, Consolas, monospace"
SANS = "'Segoe UI', -apple-system, Helvetica, Arial, sans-serif"


def colour(line):
    t = html.escape(line)
    if line.startswith("$"):
        return f"<span style='color:#7ee787'>{t}</span>"
    if line.startswith("[logged"):
        return f"<span style='color:#3fb950'>{t}</span>"
    if line.startswith("[review") or line.startswith("[error"):
        return f"<span style='color:#d29922'>{t}</span>"
    if line.startswith("[skipped") or line.startswith("[duplicate") or line.startswith("[ignored"):
        return f"<span style='color:#8b949e'>{t}</span>"
    return t


def terminal(lines, title):
    body = "\n".join(colour(x) for x in lines)
    return f"""<html><body style="margin:0;background:#0d1117"><div id="wrap" style="display:inline-block;padding:22px 28px;background:#0d1117">
<div style="font:600 15px {SANS};color:#c9d1d9;margin-bottom:12px">{html.escape(title)}</div>
<pre style="margin:0;font:13.5px/1.55 {MONO};color:#c9d1d9">{body}</pre></div></body></html>"""


def excel(rows):
    cols = ["Received", "From", "Company", "Type", "Priority", "Summary", "Due", "Meeting hold", "Status"]
    head = "".join(f"<th>{c}</th>" for c in cols)
    trs = ""
    for r in rows:
        tds = ""
        for c in cols:
            v = html.escape(str(r.get(c, "")))
            style = " style='color:#c50f1f;font-weight:600'" if c == "Priority" and v == "urgent" else ""
            tds += f"<td{style}>{v}</td>"
        trs += f"<tr>{tds}</tr>"
    return f"""<html><body style="margin:0;background:#fff"><div id="wrap" style="display:inline-block;padding:18px;background:#f3f2f1">
<div style="font:600 15px {SANS};color:#107c41;margin-bottom:10px">Excel Online: AI Inbox/Requests.xlsx, table "Requests"</div>
<table style="border-collapse:collapse;font:13px {SANS};background:#fff;max-width:1500px">
<style>th{{background:#107c41;color:#fff;text-align:left;padding:6px 9px;border:1px solid #0b5a2f}}
td{{padding:6px 9px;border:1px solid #d1d1d1;vertical-align:top}} td:nth-child(6){{max-width:430px}}</style>
<tr>{head}</tr>{trs}</table></div></body></html>"""


def hold_draft(holds, drafts):
    cards = ""
    for h in holds:
        cards += f"""<div style="border-left:5px solid #0f6cbd;background:#ebf3fc;padding:10px 14px;margin-bottom:10px;border-radius:4px">
<div style="font-weight:600">{html.escape(h['subject'])}</div>
<div style="color:#424242">{html.escape(h['start']['dateTime'][:16].replace('T', ' '))} to {html.escape(h['end']['dateTime'][11:16])} ({html.escape(h['start']['timeZone'])}), shown as {html.escape(h['showAs'])}, no invite sent</div></div>"""
    ds = ""
    for d in drafts:
        ds += f"""<div style="background:#fff;border:1px solid #d1d1d1;padding:12px 14px;margin-bottom:10px;border-radius:4px">
<div style="color:#c50f1f;font-size:12px;font-weight:600">[Draft] not sent</div><div style="font-weight:600">{html.escape(d['subject'])}</div>
<pre style="font:13px {SANS};white-space:pre-wrap;margin:6px 0 0">{html.escape(d['body']['content'])}</pre></div>"""
    return f"""<html><body style="margin:0"><div id="wrap" style="display:inline-block;width:760px;padding:18px;background:#f5f5f5;font:13.5px {SANS}">
<div style="font:600 15px {SANS};color:#0f6cbd;margin-bottom:10px">Outlook calendar: tentative holds</div>{cards}
<div style="font:600 15px {SANS};color:#0f6cbd;margin:16px 0 10px">Outlook Drafts: replies waiting for a person</div>{ds}</div></body></html>"""


def shot(page, doc, name):
    page.set_content(doc)
    page.locator("#wrap").screenshot(path=str(MEDIA / name))


def main():
    MEDIA.mkdir(exist_ok=True)
    lines = (RUN / "run.txt").read_text().rstrip().split("\n")
    rows = json.loads((RUN / "excel-rows.json").read_text())
    holds = json.loads((RUN / "calendar.json").read_text())
    drafts = json.loads((RUN / "drafts.json").read_text())
    title = "graph-assistant: real Claude Sonnet 5.5" + (", real Microsoft Graph" if "live" in RUN.name else ", Graph fake")
    with sync_playwright() as p:
        b = p.chromium.launch()
        page = b.new_page(device_scale_factor=2, viewport={"width": 1600, "height": 900})
        shot(page, terminal(lines, title), f"{PREFIX}run.png")
        shot(page, excel(rows), f"{PREFIX}excel.png")
        shot(page, hold_draft(holds, drafts), f"{PREFIX}hold-draft.png")
        b.close()


if __name__ == "__main__":
    main()
