"""
Simple local web viewer for robotics_papers_database.xlsx

Run:
    python view.py

Opens a browser tab at http://localhost:8765 showing the spreadsheet as a
searchable list of paper cards. Click a card to read the full record.
The file is re-read from disk on every page load, so editing the .xlsx
and refreshing the browser picks up the latest data.
"""

import json
import socket
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import openpyxl

EXCEL_PATH = Path(__file__).parent / "robotics_papers_database.xlsx"
PORT = 8765

# Maps the spreadsheet's header text -> a short internal field name.
# Matching is done by substring so small wording tweaks in the header
# don't break the app.
FIELD_MAP = [
    ("title", "Title"),
    ("authors", "Authors"),
    ("venue", "Journal"),
    ("date", "Published Date"),
    ("doi", "DOI"),
    ("robotType", "Robot Type"),
    ("physicalSpecs", "Physical Specs"),
    ("actuation", "Actuation"),
    ("performance", "Performance"),
    ("methodology", "Methodology"),
    ("limitations", "Limitations"),
]


def load_papers():
    wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h or "").strip() for h in rows[0]]

    col_for_field = {}
    for field, needle in FIELD_MAP:
        for i, h in enumerate(headers):
            if needle.lower() in h.lower():
                col_for_field[field] = i
                break

    papers = []
    for row in rows[1:]:
        if row is None or all(v is None for v in row):
            continue
        paper = {"headers": headers}
        for field, _ in FIELD_MAP:
            idx = col_for_field.get(field)
            value = row[idx] if idx is not None and idx < len(row) else None
            paper[field] = "" if value is None else str(value)
        papers.append(paper)
    return papers


PAGE_TEMPLATE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Robotics Papers Database</title>
<style>
  :root {
    color-scheme: light dark;
    --bg: #f7f7f8;
    --card-bg: #ffffff;
    --text: #1a1a1d;
    --muted: #6b7280;
    --border: #e5e7eb;
    --accent: #2563eb;
    --accent-bg: #eff6ff;
  }
  @media (prefers-color-scheme: dark) {
    :root {
      --bg: #16171a;
      --card-bg: #1f2023;
      --text: #f2f2f3;
      --muted: #9ca3af;
      --border: #2e3035;
      --accent: #60a5fa;
      --accent-bg: #1e293b;
    }
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    background: var(--bg);
    color: var(--text);
    font-family: -apple-system, "Segoe UI", Roboto, sans-serif;
    padding-block: 24px;
  }
  body.modal-open { overflow: hidden; }
  .wrap { max-width: 900px; margin: 0 auto; padding: 0 16px; }
  h1 { font-size: 1.4rem; margin: 0 0 4px; }
  .sub { color: var(--muted); font-size: 0.9rem; margin-bottom: 18px; }
  #search {
    width: 100%;
    padding: 10px 14px;
    font-size: 1rem;
    border: 1px solid var(--border);
    border-radius: 8px;
    background: var(--card-bg);
    color: var(--text);
    margin-bottom: 18px;
  }
  #count { font-size: 0.85rem; color: var(--muted); margin-bottom: 10px; }
  .card {
    background: var(--card-bg);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 14px 16px;
    margin-bottom: 12px;
    cursor: pointer;
    transition: border-color .15s;
  }
  .card:hover { border-color: var(--accent); }
  .card-title { font-weight: 600; font-size: 1.02rem; margin-bottom: 4px; }
  .card-meta { font-size: 0.85rem; color: var(--muted); margin-bottom: 6px; }
  .card-snippet {
    font-size: 0.88rem;
    color: var(--text);
    opacity: 0.85;
    display: -webkit-box;
    -webkit-line-clamp: 2;
    -webkit-box-orient: vertical;
    overflow: hidden;
  }
  .badge {
    display: inline-block;
    background: var(--accent-bg);
    color: var(--accent);
    border-radius: 999px;
    padding: 2px 9px;
    font-size: 0.78rem;
    margin-right: 6px;
  }
  #empty { color: var(--muted); padding: 20px 0; display: none; }

  .overlay {
    display: none;
    position: fixed; inset: 0;
    background: rgba(0,0,0,0.45);
    align-items: flex-start;
    justify-content: center;
    padding: 30px 16px;
    overflow-y: auto;
    z-index: 10;
  }
  .overlay.open { display: flex; }
  .modal {
    --font-scale: 1;
    background: var(--card-bg);
    color: var(--text);
    max-width: 720px;
    width: 100%;
    border-radius: 12px;
    padding: 0 26px 30px;
    position: relative;
  }
  .modal h2 { margin: 0 26px 10px 0; font-size: calc(1.2rem * var(--font-scale)); }
  .modal .card-meta { margin-bottom: 14px; font-size: calc(0.85rem * var(--font-scale)); }
  .field { margin-bottom: 16px; }
  .field-label {
    font-size: calc(0.75rem * var(--font-scale));
    text-transform: uppercase;
    letter-spacing: 0.04em;
    color: var(--accent);
    font-weight: 600;
    margin-bottom: 4px;
  }
  .field-value {
    font-size: calc(0.92rem * var(--font-scale));
    line-height: 1.5;
    white-space: pre-wrap;
  }
  .field-value a { color: var(--accent); }
  .modal-controls {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    padding: 16px 0;
    margin-bottom: 16px;
    border-bottom: 1px solid var(--border);
  }
  .text-size-btn {
    background: none;
    border: 1px solid var(--border);
    border-radius: 6px;
    cursor: pointer;
    color: var(--muted);
    line-height: 1;
    font-size: 0.8rem;
    font-weight: 600;
    padding: 8px 12px;
  }
  .text-size-btn:hover { color: var(--text); border-color: var(--accent); }
  .close-btn {
    position: absolute;
    top: 14px; right: 14px;
    background: none;
    border: none;
    font-size: 1.3rem;
    cursor: pointer;
    color: var(--muted);
    line-height: 1;
    padding: 4px 8px;
  }
  .close-btn:hover { color: var(--text); }

  .scroll-top-btn {
    position: fixed;
    right: 24px;
    bottom: 24px;
    width: 46px;
    height: 46px;
    border-radius: 50%;
    background: var(--accent);
    color: #fff;
    border: none;
    font-size: 1.3rem;
    cursor: pointer;
    box-shadow: 0 2px 10px rgba(0,0,0,0.3);
    align-items: center;
    justify-content: center;
    display: none;
    z-index: 5;
  }
  .scroll-top-btn.show { display: flex; }
  .scroll-top-btn:hover { filter: brightness(1.1); }
</style>
</head>
<body>
<div class="wrap">
  <h1>Robotics Papers Database</h1>
  <div class="sub" id="sub"></div>
  <input id="search" type="text" placeholder="Search title, authors, robot type, methodology…">
  <div id="count"></div>
  <div id="list"></div>
  <div id="empty">No papers match your search.</div>
</div>

<button class="scroll-top-btn" id="scrollTopBtn" title="Scroll to top">&uarr;</button>

<div class="overlay" id="overlay">
  <div class="modal" id="modal">
    <button class="close-btn" id="closeBtn" title="Close">&times;</button>
    <div class="modal-controls">
      <button class="text-size-btn" id="textSmaller" title="Decrease text size">A&minus;</button>
      <button class="text-size-btn" id="textReset" title="Reset text size">Reset</button>
      <button class="text-size-btn" id="textBigger" title="Increase text size">A+</button>
    </div>
    <div id="modalBody"></div>
  </div>
</div>

<script id="data" type="application/json">__DATA__</script>
<script>
const papers = JSON.parse(document.getElementById('data').textContent);

const FIELDS = [
  ["robotType", "Robot Type & Locomotion Mechanism"],
  ["physicalSpecs", "Physical Specs"],
  ["actuation", "Actuation & Control System"],
  ["performance", "Performance Metrics"],
  ["methodology", "Methodology / Robot Summary"],
  ["limitations", "Limitations & Future Work"],
];

document.getElementById('sub').textContent = papers.length + " paper" + (papers.length === 1 ? "" : "s") + " loaded from robotics_papers_database.xlsx";

const listEl = document.getElementById('list');
const emptyEl = document.getElementById('empty');
const countEl = document.getElementById('count');
const searchEl = document.getElementById('search');
const overlay = document.getElementById('overlay');
const modalBody = document.getElementById('modalBody');

function escapeHtml(s) {
  return (s || "").replace(/[&<>]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;"}[c]));
}

function linkifyDoi(doi) {
  const trimmed = (doi || "").trim();
  if (/^https?:\\/\\//i.test(trimmed)) {
    return `<a href="${escapeHtml(trimmed)}" target="_blank" rel="noopener">${escapeHtml(trimmed)}</a>`;
  }
  return escapeHtml(trimmed);
}

function render(items) {
  listEl.innerHTML = "";
  emptyEl.style.display = items.length ? "none" : "block";
  countEl.textContent = items.length + " shown";
  items.forEach(p => {
    const card = document.createElement('div');
    card.className = 'card';
    card.innerHTML = `
      <div class="card-title">${escapeHtml(p.title) || "(Untitled)"}</div>
      <div class="card-meta">
        ${p.venue ? `<span class="badge">${escapeHtml(p.venue)}</span>` : ""}
        ${p.date ? escapeHtml(p.date) : ""}
      </div>
      <div class="card-meta">${escapeHtml(p.authors)}</div>
      <div class="card-snippet">${escapeHtml(p.robotType)}</div>
    `;
    card.addEventListener('click', () => openModal(p));
    listEl.appendChild(card);
  });
}

function openModal(p) {
  let html = `<h2>${escapeHtml(p.title) || "(Untitled)"}</h2>`;
  html += `<div class="card-meta">${escapeHtml(p.authors)}</div>`;
  html += `<div class="card-meta">${escapeHtml(p.venue)}${p.venue && p.date ? " &middot; " : ""}${escapeHtml(p.date)}</div>`;
  if (p.doi) {
    html += `<div class="field"><div class="field-label">DOI / URL</div><div class="field-value">${linkifyDoi(p.doi)}</div></div>`;
  }
  FIELDS.forEach(([key, label]) => {
    if (p[key]) {
      html += `<div class="field"><div class="field-label">${escapeHtml(label)}</div><div class="field-value">${escapeHtml(p[key])}</div></div>`;
    }
  });
  modalBody.innerHTML = html;
  overlay.classList.add('open');
  document.body.classList.add('modal-open');
}

function closeModal() {
  overlay.classList.remove('open');
  document.body.classList.remove('modal-open');
}

document.getElementById('closeBtn').addEventListener('click', closeModal);
overlay.addEventListener('click', e => { if (e.target === overlay) closeModal(); });
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });

const modalEl = document.getElementById('modal');
const FONT_SCALE_MIN = 0.75;
const FONT_SCALE_MAX = 2.0;
const FONT_SCALE_STEP = 0.1;
let fontScale = parseFloat(localStorage.getItem('paperFontScale')) || 1;

function applyFontScale() {
  fontScale = Math.min(FONT_SCALE_MAX, Math.max(FONT_SCALE_MIN, fontScale));
  modalEl.style.setProperty('--font-scale', fontScale);
  try { localStorage.setItem('paperFontScale', fontScale); } catch (e) {}
}
applyFontScale();

document.getElementById('textBigger').addEventListener('click', () => { fontScale += FONT_SCALE_STEP; applyFontScale(); });
document.getElementById('textSmaller').addEventListener('click', () => { fontScale -= FONT_SCALE_STEP; applyFontScale(); });
document.getElementById('textReset').addEventListener('click', () => { fontScale = 1; applyFontScale(); });

const scrollTopBtn = document.getElementById('scrollTopBtn');
window.addEventListener('scroll', () => {
  scrollTopBtn.classList.toggle('show', window.scrollY > 300);
});
scrollTopBtn.addEventListener('click', () => window.scrollTo({ top: 0, behavior: 'smooth' }));

searchEl.addEventListener('input', () => {
  const q = searchEl.value.trim().toLowerCase();
  if (!q) { render(papers); return; }
  const filtered = papers.filter(p =>
    Object.values(p).some(v => typeof v === "string" && v.toLowerCase().includes(q))
  );
  render(filtered);
});

render(papers);
</script>
</body>
</html>
"""


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass  # keep the console quiet

    def do_GET(self):
        if self.path not in ("/", "/index.html"):
            self.send_response(404)
            self.end_headers()
            return
        try:
            papers = load_papers()
        except FileNotFoundError:
            self.send_response(500)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.end_headers()
            self.wfile.write(f"Excel file not found: {EXCEL_PATH}".encode("utf-8"))
            return

        data_json = json.dumps(papers, ensure_ascii=False).replace("</", "<\\/")
        html = PAGE_TEMPLATE.replace("__DATA__", data_json)

        body = html.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def find_free_port(start):
    port = start
    while port < start + 50:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port
        port += 1
    return start


def main():
    port = find_free_port(PORT)
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://127.0.0.1:{port}/"
    print(f"Serving {EXCEL_PATH.name} at {url}  (Ctrl+C to stop)")
    threading.Timer(0.4, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping.")
        server.shutdown()


if __name__ == "__main__":
    main()
