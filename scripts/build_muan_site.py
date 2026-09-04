#!/usr/bin/env python3
"""Build a static product lookup site from the MUAN WPS-exported workbook.

The workbook stores product photos as WPS DISPIMG cell images. This parser
resolves those IDs through xl/cellimages.xml, optimizes the photos, and turns
the merged size-table cells into small PNG charts for the web result page.
"""

from __future__ import annotations

import html
import json
import posixpath
import re
import shutil
import sys
import zipfile
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET

from PIL import Image, ImageDraw, ImageFont

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ETC_NS = "http://www.wps.cn/officeDocument/2017/etCustomData"
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
DRAW_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"m": MAIN_NS, "r": REL_NS, "etc": ETC_NS, "xdr": XDR_NS, "a": DRAW_NS}

IMAGE_EDGE = 1000
IMAGE_QUALITY = 78
FONT_PATH = "/System/Library/Fonts/STHeiti Medium.ttc"
SIZE_HEADER_LABELS = {"尺码", "尺寸", "尺寸表", "码数", "尺码（高个子）", "小个子尺寸"}


def col_index(ref: str) -> int:
    letters = re.match(r"[A-Z]+", ref or "")
    if not letters:
        return -1
    out = 0
    for char in letters.group(0):
        out = out * 26 + ord(char) - 64
    return out - 1


def shared_strings(zf: zipfile.ZipFile) -> list[str]:
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    return ["".join(t.text or "" for t in node.findall(".//m:t", NS))
            for node in root.findall("m:si", NS)]


def cell_value(cell: ET.Element, strings: list[str]) -> str:
    value = cell.find("m:v", NS)
    raw = value.text if value is not None else ""
    if cell.attrib.get("t") == "s" and raw:
        return strings[int(raw)]
    return raw or ""


def read_rows(zf: zipfile.ZipFile, strings: list[str]) -> dict[int, dict[int, str]]:
    root = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
    rows: dict[int, dict[int, str]] = {}
    for row in root.findall("m:sheetData/m:row", NS):
        row_no = int(row.attrib["r"]) - 1
        values: dict[int, str] = {}
        for cell in row.findall("m:c", NS):
            index = col_index(cell.attrib.get("r", ""))
            if index >= 0:
                values[index] = cell_value(cell, strings)
                formula = cell.find("m:f", NS)
                if formula is not None and formula.text:
                    values[index] = f"__FORMULA__{formula.text}"
        rows[row_no] = values
    return rows


def rels_map(zf: zipfile.ZipFile, path: str) -> dict[str, str]:
    root = ET.fromstring(zf.read(path))
    return {node.attrib["Id"]: node.attrib["Target"] for node in root}


def cell_image_map(zf: zipfile.ZipFile) -> dict[str, str]:
    root = ET.fromstring(zf.read("xl/cellimages.xml"))
    rels = rels_map(zf, "xl/_rels/cellimages.xml.rels")
    result: dict[str, str] = {}
    for image in root.findall("etc:cellImage", NS):
        name = image.find(".//xdr:cNvPr", NS)
        blip = image.find(".//a:blip", NS)
        if name is None or blip is None:
            continue
        image_id = name.attrib.get("name", "")
        rid = blip.attrib.get(f"{{{REL_NS}}}embed", "")
        target = rels.get(rid)
        if image_id and target:
            result[image_id] = posixpath.normpath(posixpath.join("xl", target))
    return result


def image_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    try:
        return ImageFont.truetype(FONT_PATH, size=size, index=0)
    except OSError:
        return ImageFont.load_default()


def optimize_image(data: bytes) -> tuple[bytes, str]:
    with Image.open(BytesIO(data)) as source:
        source.load()
        image = source.convert("RGB")
        scale = min(1.0, IMAGE_EDGE / max(image.size))
        if scale < 1:
            image = image.resize((round(image.width * scale), round(image.height * scale)), Image.Resampling.LANCZOS)
        output = BytesIO()
        image.save(output, format="JPEG", quality=IMAGE_QUALITY, optimize=True, progressive=True)
        return output.getvalue(), ".jpg"


def text_value(value: str) -> str:
    if not value or value.startswith("__FORMULA__"):
        return ""
    return value.strip()


def make_size_chart(code: str, headers: list[str], measurements: list[tuple[str, list[str]]], output: Path) -> bool:
    if not headers or not measurements:
        return False
    label_font = image_font(22)
    value_font = image_font(21)
    title_font = image_font(25)
    label_width = 310
    col_width = 116
    row_height = 48
    width = label_width + col_width * len(headers) + 28
    height = 74 + row_height * (len(measurements) + 1) + 24
    canvas = Image.new("RGB", (width, height), "#ffffff")
    draw = ImageDraw.Draw(canvas)
    draw.rounded_rectangle((8, 8, width - 8, height - 8), radius=12, outline="#d8dee8", width=2, fill="#ffffff")
    draw.text((24, 20), f"{code} 尺码表", fill="#17233b", font=title_font)
    top = 70
    draw.rectangle((14, top, width - 14, top + row_height), fill="#edf3f8")
    draw.text((26, top + 12), "测量项目", fill="#31425e", font=label_font)
    for index, header in enumerate(headers):
        x = label_width + 14 + index * col_width
        draw.text((x + 28, top + 12), header, fill="#31425e", font=value_font)
    for row_index, (label, values) in enumerate(measurements, start=1):
        y = top + row_index * row_height
        if row_index % 2 == 0:
            draw.rectangle((14, y, width - 14, y + row_height), fill="#fafbfd")
        draw.text((26, y + 12), label[:18], fill="#334155", font=label_font)
        for col_index_, value in enumerate(values):
            x = label_width + 14 + col_index_ * col_width
            draw.text((x + 22, y + 12), value, fill="#17233b", font=value_font)
    output.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(output, format="PNG", optimize=True)
    return True


def extract_size_table(rows: dict[int, dict[int, str]], start: int, end: int) -> tuple[list[str], list[tuple[str, list[str]]]]:
    """Find the first usable size table inside one product block.

    WPS exports use several labels (尺寸, 尺码表, etc.) and sometimes put a
    section title on one row and the actual size header on the next row.
    """
    candidates: list[int] = []
    for row_no in range(start, end):
        row = rows.get(row_no, {})
        label = text_value(row.get(8, ""))
        values = [text_value(row.get(col, "")) for col in range(9, 14)]
        if label in SIZE_HEADER_LABELS or label.startswith(("尺码", "尺寸", "码数")):
            if any(values):
                candidates.append(row_no)
        elif not label and any(values) and row_no + 1 < end:
            next_row = rows.get(row_no + 1, {})
            next_label = text_value(next_row.get(8, ""))
            next_values = [text_value(next_row.get(col, "")) for col in range(9, 14)]
            if next_label and any(re.search(r"\d", value) for value in next_values):
                candidates.append(row_no)

    for header_row in candidates:
        raw_headers = [text_value(rows[header_row].get(col, "")) for col in range(9, 14)]
        while raw_headers and not raw_headers[-1]:
            raw_headers.pop()
        if not raw_headers:
            continue
        measurements: list[tuple[str, list[str]]] = []
        for row_no in range(header_row + 1, end):
            row = rows.get(row_no, {})
            label = text_value(row.get(8, ""))
            values = [text_value(row.get(col, "")) for col in range(9, 14)]
            if not label and not any(values):
                if measurements:
                    break
                continue
            if label in SIZE_HEADER_LABELS or label.startswith(("尺码", "尺寸", "码数")):
                break
            if label and any(values):
                measurements.append((label, values[:len(raw_headers)]))
            elif measurements:
                break
        if measurements:
            return raw_headers, measurements

    # Some one-size products have measurement rows but no size names at all.
    # Keep those source values visible with a neutral header instead of dropping
    # the table or inventing a size label.
    fallback: list[tuple[str, list[str]]] = []
    max_columns = 0
    for row_no in range(start, end):
        row = rows.get(row_no, {})
        label = text_value(row.get(8, ""))
        values = [text_value(row.get(col, "")) for col in range(9, 14)]
        if label and any(values) and not (label in SIZE_HEADER_LABELS or label.startswith(("尺码", "尺寸", "码数"))):
            fallback.append((label, values))
            max_columns = max(max_columns, max((index + 1 for index, value in enumerate(values) if value), default=0))
        elif fallback and not label and not any(values):
            break
    if fallback and max_columns:
        return ["数值"] if max_columns == 1 else [f"规格{index + 1}" for index in range(max_columns)], [
            (label, values[:max_columns]) for label, values in fallback
        ]

    return [], []


def html_escape(value: object) -> str:
    return html.escape(str(value or ""), quote=True)


def build_page(products: list[dict], title: str, output: Path) -> None:
    product_json = json.dumps(products, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c")
    page_title = html_escape(title or "MUAN 商品资料查询")
    page = f'''<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{page_title} · 货号查询</title>
<style>
:root{{--ink:#17233b;--muted:#6b778c;--line:#dbe2ea;--paper:#fff;--canvas:#f4f7fa;--accent:#0e7490;--accent2:#e76f51;--soft:#e8f3f5}}
*{{box-sizing:border-box}} body{{margin:0;background:var(--canvas);color:var(--ink);font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC",sans-serif}}
button,input{{font:inherit}} button{{cursor:pointer}} .top{{background:linear-gradient(120deg,#102a43,#0e7490);color:#fff;padding:28px 20px 30px}}
.top-inner{{max-width:1180px;margin:0 auto}} .eyebrow{{font-size:12px;letter-spacing:.12em;opacity:.72;text-transform:uppercase}} h1{{font-size:clamp(25px,4vw,42px);margin:8px 0 20px;letter-spacing:0}}
.search{{display:flex;gap:10px;max-width:760px}} .search input{{flex:1;min-width:0;border:0;border-radius:8px;padding:14px 16px;color:var(--ink);background:#fff;outline:none;box-shadow:0 4px 12px #081c2c33}}
.search button{{border:0;border-radius:8px;padding:0 22px;background:var(--accent2);color:#fff;font-weight:700}} .wrap{{max-width:1180px;margin:0 auto;padding:24px 20px 60px}}
.toolbar{{display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:18px;color:var(--muted);font-size:14px}} .toolbar strong{{color:var(--ink)}}
.reset{{border:1px solid var(--line);background:var(--paper);color:var(--ink);border-radius:7px;padding:8px 12px}} .grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:16px}}
.card{{background:var(--paper);border:1px solid var(--line);border-radius:10px;overflow:hidden;box-shadow:0 4px 16px #20364d0b;transition:transform .18s,box-shadow .18s}} .card:hover{{transform:translateY(-2px);box-shadow:0 10px 24px #20364d18}}
.thumb{{height:220px;background:#edf1f4;display:flex;align-items:center;justify-content:center}} .thumb img{{width:100%;height:100%;object-fit:contain}} .card-body{{padding:15px 16px 17px}} .code{{font-size:18px;font-weight:750;letter-spacing:.02em}} .meta{{display:flex;gap:8px;flex-wrap:wrap;margin-top:9px;color:var(--muted);font-size:13px}} .pill{{background:var(--soft);color:var(--accent);padding:4px 8px;border-radius:999px}}
.empty{{grid-column:1/-1;text-align:center;padding:72px 20px;color:var(--muted);background:var(--paper);border:1px dashed var(--line);border-radius:10px}}
.shade{{display:none;position:fixed;inset:0;background:#081c2c99;z-index:10;padding:18px;align-items:center;justify-content:center}} .shade.open{{display:flex}} .modal{{width:min(960px,100%);max-height:94vh;overflow:auto;background:var(--paper);border-radius:12px}}
.modal-head{{display:flex;align-items:center;justify-content:space-between;padding:18px 22px;border-bottom:1px solid var(--line);position:sticky;top:0;background:#fff;z-index:2}} .modal-head h2{{margin:0;font-size:22px}} .close{{border:0;background:transparent;font-size:28px;color:var(--muted);line-height:1}}
.detail{{display:grid;grid-template-columns:minmax(250px,330px) 1fr;gap:24px;padding:22px}} .hero{{border:1px solid var(--line);border-radius:9px;background:#f1f4f6;display:flex;align-items:center;justify-content:center;min-height:250px}} .hero img{{width:100%;height:100%;max-height:430px;object-fit:contain}}
.section{{margin-bottom:22px}} .section h3{{font-size:14px;color:var(--accent);margin:0 0 10px;border-bottom:2px solid var(--soft);padding-bottom:8px}} .facts{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:9px}} .fact{{background:#f7f9fb;border-radius:7px;padding:10px 12px}} .fact small{{display:block;color:var(--muted);font-size:11px;margin-bottom:4px}} .fact div{{white-space:pre-wrap;line-height:1.5;word-break:break-word}}
.long{{white-space:pre-wrap;line-height:1.65;color:#334155;background:#f7f9fb;border-radius:7px;padding:12px}} .chart{{display:block;width:100%;border:1px solid var(--line);border-radius:8px;background:#fff;cursor:zoom-in}} .chart-note{{margin-top:8px;color:var(--muted);font-size:12px}} .actions{{display:flex;justify-content:flex-end;gap:10px;margin-top:12px;flex-wrap:wrap}} .action{{border:1px solid var(--line);border-radius:7px;padding:9px 13px;background:#fff;color:var(--ink)}} .action.primary{{background:var(--accent);border-color:var(--accent);color:#fff}} .chart-overlay{{display:none;position:fixed;inset:0;background:#081c2ce6;z-index:20;padding:24px;align-items:center;justify-content:center}} .chart-overlay.open{{display:flex}} .chart-overlay img{{max-width:100%;max-height:100%;object-fit:contain;background:#fff;border-radius:8px}} .chart-overlay-close{{position:absolute;top:12px;right:18px;border:0;background:transparent;color:#fff;font-size:34px;line-height:1;cursor:pointer}}
.more-wrap{{display:flex;justify-content:center;margin:24px 0 4px}} .more-wrap button{{border:1px solid var(--line);border-radius:7px;padding:10px 16px;background:var(--paper);color:var(--ink)}}
@media(max-width:680px){{.top{{padding:22px 16px 24px}} .wrap{{padding:18px 16px 40px}} .search button{{padding:0 16px}} .detail{{grid-template-columns:1fr;padding:16px;gap:16px}} .hero{{min-height:220px}} .facts{{grid-template-columns:1fr}} .modal-head{{padding:15px 16px}}}}
</style>
</head>
<body>
<header class="top"><div class="top-inner"><div class="eyebrow">MUAN PRODUCT LIBRARY</div><h1>{page_title}</h1><div class="search"><input id="query" placeholder="输入货号，例如 660002" autocomplete="off"><button id="search">查询</button></div></div></header>
<main class="wrap"><div class="toolbar"><div id="stats"></div><button class="reset" id="reset">显示全部</button></div><section class="grid" id="grid"></section><div class="more-wrap"><button id="more" hidden>加载更多</button></div></main>
<div class="shade" id="shade"><article class="modal"><div class="modal-head"><h2 id="modal-title"></h2><button class="close" id="close" aria-label="关闭">×</button></div><div id="modal-body"></div></article></div>
<div class="chart-overlay" id="chart-overlay" role="dialog" aria-modal="true" aria-label="放大尺码表"><button class="chart-overlay-close" id="chart-overlay-close" aria-label="关闭放大图">×</button><img id="chart-preview" alt="放大尺码表"></div>
<script>
const products = {product_json};
const $ = (s) => document.querySelector(s);
const esc = (v) => String(v ?? "").replace(/[&<>"']/g, c => c === "&" ? "&amp;" : c === "<" ? "&lt;" : c === ">" ? "&gt;" : c === String.fromCharCode(34) ? "&quot;" : "&#39;");
const nl = (v) => esc(v).replace(/\\n/g,"<br>");
const PAGE_SIZE = 120;
let activeList = products;
let visibleCount = PAGE_SIZE;
function render(list, resetPage = true) {{
  if (resetPage) {{ activeList = list; visibleCount = PAGE_SIZE; }}
  const shown = activeList.slice(0, visibleCount);
  $("#stats").innerHTML = `共 <strong>${{products.length}}</strong> 个商品，当前结果 <strong>${{activeList.length}}</strong> 个，已显示 <strong>${{shown.length}}</strong> 个`;
  $("#grid").innerHTML = shown.length ? shown.map(p => `<article class="card" data-index="${{products.indexOf(p)}}"><div class="thumb">${{p.product_image ? `<img src="${{esc(p.product_image)}}" loading="lazy" alt="${{esc(p.code)}}">` : "<span>暂无图片</span>"}}</div><div class="card-body"><div class="code">${{esc(p.code)}}</div><div class="meta"><span class="pill">${{esc(p.color || "颜色待补")}}</span><span>${{esc(p.material || "材质待补")}}</span></div></div></article>`).join("") : '<div class="empty">没有找到匹配的货号</div>';
  const more = $("#more");
  more.hidden = shown.length >= activeList.length;
  more.textContent = `加载更多（剩余 ${{Math.max(0, activeList.length - shown.length)}} 条）`;
  document.querySelectorAll(".card").forEach(card => card.addEventListener("click", () => show(Number(card.dataset.index))));
}}
function show(i) {{
  const p = products[i]; $("#modal-title").textContent = p.code || "商品详情";
  const facts = [["颜色",p.color],["面料及成分",p.material],["订货日期",p.order_date],["订货数量",p.order_qty],["裁剪数量",p.cut_qty],["出货日期",p.ship_date]].filter(x => x[1]);
  let body = `<div class="detail"><div><div class="hero">${{p.product_image ? `<img src="${{esc(p.product_image)}}" alt="${{esc(p.code)}}">` : "<span>暂无商品图</span>"}}</div><div class="actions">${{p.product_image ? `<a class="action" href="${{esc(p.product_image)}}" target="_blank" rel="noreferrer">查看原图</a>` : ""}}<button class="action primary" id="copy">复制资料</button></div></div><div>`;
  body += `<section class="section"><h3>基础资料</h3><div class="facts">${{facts.map(x => `<div class="fact"><small>${{esc(x[0])}}</small><div>${{nl(x[1])}}</div></div>`).join("")}}</div></section>`;
  if (p.size_chart) body += `<section class="section"><h3>尺码表</h3><img class="chart" id="chart" src="${{esc(p.size_chart)}}" alt="${{esc(p.code)}} 尺码表" loading="eager" tabindex="0"><div class="chart-note">点击图片可放大查看</div><div class="actions"><button class="action" id="copy-chart">复制尺码表图片</button><a class="action" href="${{esc(p.size_chart)}}" download="${{esc(p.code)}}-尺码表.png">下载尺码表图片</a></div></section>`;
  if (p.highlight) body += `<section class="section"><h3>卖点 / 备注</h3><div class="long">${{nl(p.highlight)}}</div></section>`;
  body += `</div></div>`; $("#modal-body").innerHTML = body; $("#shade").classList.add("open");
  $("#copy").addEventListener("click", async () => {{ const text = [p.code && `货号：${{p.code}}`,p.color && `颜色：${{p.color}}`,p.material && `面料及成分：${{p.material}}`,p.order_date && `订货日期：${{p.order_date}}`,p.order_qty && `订货数量：${{p.order_qty}}`,p.cut_qty && `裁剪数量：${{p.cut_qty}}`,p.ship_date && `出货日期：${{p.ship_date}}`,p.highlight && `卖点/备注：\\n${{p.highlight}}`].filter(Boolean).join("\\n"); try {{ await navigator.clipboard.writeText(text); $("#copy").textContent = "已复制"; }} catch {{ $("#copy").textContent = "请手动复制"; }} }});
  if (p.size_chart) {{
    const chart = $("#chart");
    const openChart = () => {{ $("#chart-preview").src = p.size_chart; $("#chart-overlay").classList.add("open"); }};
    chart.addEventListener("click", openChart);
    chart.addEventListener("keydown", e => {{ if (e.key === "Enter" || e.key === " ") {{ e.preventDefault(); openChart(); }} }});
    $("#copy-chart").addEventListener("click", async () => {{
      const button = $("#copy-chart");
      try {{
        const blob = await fetch(p.size_chart).then(r => r.blob());
        await navigator.clipboard.write([new ClipboardItem({{"image/png": blob}})]);
        button.textContent = "已复制图片";
      }} catch {{ button.textContent = "请长按图片复制"; }}
    }});
  }}
}}
function filter(syncUrl = true) {{
  const raw = $("#query").value.trim();
  const q = raw.toLowerCase();
  if (syncUrl) history.replaceState(null, "", raw ? `?code=${{encodeURIComponent(raw)}}` : location.pathname);
  if (!q) return render(products);
  const exact = products.filter(p => String(p.code || "").trim().toLowerCase() === q);
  render(exact.length ? exact : products.filter(p => [p.code,p.color,p.material,p.highlight].join(" ").toLowerCase().includes(q)));
}}
$("#search").addEventListener("click", () => filter()); $("#query").addEventListener("keydown", e => {{ if (e.key === "Enter") filter(); }}); $("#reset").addEventListener("click", () => {{ $("#query").value = ""; filter(); }}); $("#more").addEventListener("click", () => {{ visibleCount += PAGE_SIZE; render(activeList, false); }}); $("#close").addEventListener("click", () => $("#shade").classList.remove("open")); $("#shade").addEventListener("click", e => {{ if (e.target === $("#shade")) $("#shade").classList.remove("open"); }}); $("#chart-overlay-close").addEventListener("click", () => $("#chart-overlay").classList.remove("open")); $("#chart-overlay").addEventListener("click", e => {{ if (e.target === $("#chart-overlay")) $("#chart-overlay").classList.remove("open"); }}); document.addEventListener("keydown", e => {{ if (e.key === "Escape") {{ $("#chart-overlay").classList.remove("open"); $("#shade").classList.remove("open"); }} }}); const initialQuery = new URLSearchParams(location.search).get("code") || ""; if (initialQuery) {{ $("#query").value = initialQuery; filter(false); }} else render(products);
</script>
</body>
</html>'''
    output.write_text(page, encoding="utf-8")


def main() -> None:
    if len(sys.argv) != 3:
        raise SystemExit("Usage: build_muan_site.py source.xlsx output_dir")
    source = Path(sys.argv[1]).expanduser().resolve()
    destination = Path(sys.argv[2]).expanduser().resolve()
    image_dir = destination / "data" / "images"
    chart_dir = destination / "data" / "charts"
    image_dir.mkdir(parents=True, exist_ok=True)
    chart_dir.mkdir(parents=True, exist_ok=True)
    for old in list(image_dir.glob("*")) + list(chart_dir.glob("*")):
        if old.is_file():
            old.unlink()

    with zipfile.ZipFile(source) as zf:
        strings = shared_strings(zf)
        rows = read_rows(zf, strings)
        image_map = cell_image_map(zf)
        media_cache: dict[str, bytes] = {}
        starts = []
        for row_no, values in rows.items():
            serial = text_value(values.get(0, ""))
            code = text_value(values.get(1, ""))
            image_formula = values.get(2, "")
            if serial.isdigit() and code and "DISPIMG" in image_formula.upper():
                starts.append(row_no)
        starts.sort()
        products: list[dict] = []
        title = text_value(rows.get(0, {}).get(0, "")) or "MUAN 商品资料"
        for index, start in enumerate(starts):
            end = starts[index + 1] if index + 1 < len(starts) else max(rows) + 1
            base = rows[start]
            code = text_value(base.get(1, ""))
            if not code:
                continue
            product = {
                "serial": text_value(base.get(0, "")),
                "code": code,
                "color": text_value(base.get(4, "")),
                "material": text_value(base.get(14, "")),
                "highlight": text_value(base.get(15, "")),
                "order_date": text_value(base.get(3, "")),
                "order_qty": text_value(base.get(5, "")),
                "cut_qty": text_value(base.get(6, "")),
                "ship_date": text_value(base.get(7, "")),
                "product_image": "",
                "size_chart": "",
            }
            image_id_match = re.search(r"DISPIMG\(\\?\"([^\"\\]+)", base.get(2, ""), re.I)
            if image_id_match:
                media_path = image_map.get(image_id_match.group(1))
                if media_path:
                    if media_path not in media_cache:
                        media_cache[media_path] = zf.read(media_path)
                    try:
                        optimized, extension = optimize_image(media_cache[media_path])
                        image_name = f"p{index + 1:04d}_{re.sub(r'[^A-Za-z0-9_-]+', '_', code)}_product{extension}"
                        (image_dir / image_name).write_bytes(optimized)
                        product["product_image"] = f"data/images/{image_name}"
                    except Exception:
                        pass

            headers, measurements = extract_size_table(rows, start, end)
            chart_name = f"p{index + 1:04d}_{re.sub(r'[^A-Za-z0-9_-]+', '_', code)}_size.png"
            chart_path = chart_dir / chart_name
            if make_size_chart(code, headers, measurements, chart_path):
                product["size_chart"] = f"data/charts/{chart_name}"
            products.append(product)

    build_page(products, title, destination / "index.html")
    (destination / "README.md").write_text(
        "# MUAN 商品查询\n\n输入货号即可查看商品资料，尺码表支持图片查看和下载。\n\n本页面默认不公开成本等内部字段。\n",
        encoding="utf-8",
    )
    total_image_bytes = sum(path.stat().st_size for path in image_dir.glob("*"))
    print(json.dumps({"products": len(products), "images": len(list(image_dir.glob("*"))), "charts": len(list(chart_dir.glob("*"))), "image_bytes": total_image_bytes}, ensure_ascii=False))


if __name__ == "__main__":
    main()
