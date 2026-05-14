#!/usr/bin/env python3
import base64
import json
import posixpath
import re
import shutil
import struct
import sys
import zipfile
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET

from PIL import Image

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "officeRel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}

DATA_FIELDS = {
    "款式编码": "code",
    "状态分类": "category",
    "材质": "material",
    "组合": "group",
    "商品等级": "level",
    "属性": "gender",
    "季节": "season",
    "在仓数量": "stock",
    "核心卖点": "highlight",
    "检测卖点": "detect",
    "包装方式": "package",
    "包装款式编码": "package_code",
    "包装辅料明细": "package_detail",
}

VISIBLE_SHEETS = {"保暖", "冰丝内裤", "棉+莫代尔内裤", "袜子", "家居服"}
MAX_IMAGE_EDGE = 1600
JPEG_QUALITY = 86


def col_name_to_index(name):
    index = 0
    for char in name:
        index = index * 26 + ord(char.upper()) - ord("A") + 1
    return index - 1


def cell_ref_to_pos(ref):
    match = re.match(r"([A-Z]+)(\d+)", ref)
    if not match:
        return None
    return int(match.group(2)) - 1, col_name_to_index(match.group(1))


def rels_map(zf, path):
    if path not in zf.namelist():
        return {}
    root = ET.fromstring(zf.read(path))
    return {rel.attrib["Id"]: rel.attrib["Target"] for rel in root}


def read_shared_strings(zf):
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings = []
    for si in root.findall("main:si", NS):
        parts = []
        for t in si.findall(".//main:t", NS):
            parts.append(t.text or "")
        strings.append("".join(parts))
    return strings


def read_cell_value(cell, shared_strings):
    cell_type = cell.attrib.get("t")
    value = cell.find("main:v", NS)
    if value is None:
        inline = cell.find("main:is/main:t", NS)
        return inline.text if inline is not None and inline.text is not None else ""
    text = value.text or ""
    if cell_type == "s":
        return shared_strings[int(text)]
    return text


def parse_workbook_sheets(zf):
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    wb_rels = rels_map(zf, "xl/_rels/workbook.xml.rels")
    sheets = []
    for sheet in workbook.findall("main:sheets/main:sheet", NS):
        name = sheet.attrib["name"]
        state = sheet.attrib.get("state", "visible")
        rid = sheet.attrib[f"{{{NS['officeRel']}}}id"]
        target = wb_rels[rid]
        path = posixpath.normpath(posixpath.join("xl", target))
        sheets.append({"name": name, "state": state, "path": path})
    return sheets


def sheet_drawing_path(zf, sheet_path, sheet_root):
    drawing = sheet_root.find("main:drawing", NS)
    if drawing is None:
        return None
    rid = drawing.attrib[f"{{{NS['officeRel']}}}id"]
    rel_path = posixpath.join(posixpath.dirname(sheet_path), "_rels", posixpath.basename(sheet_path) + ".rels")
    target = rels_map(zf, rel_path)[rid]
    return posixpath.normpath(posixpath.join(posixpath.dirname(sheet_path), target))


def image_size(data):
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return struct.unpack(">II", data[16:24])
    if data[:2] == b"\xff\xd8":
        idx = 2
        while idx < len(data) - 9:
            if data[idx] != 0xFF:
                idx += 1
                continue
            marker = data[idx + 1]
            if 0xC0 <= marker <= 0xC3:
                height, width = struct.unpack(">HH", data[idx + 5 : idx + 9])
                return width, height
            length = struct.unpack(">H", data[idx + 2 : idx + 4])[0]
            idx += 2 + length
    try:
        with Image.open(BytesIO(data)) as img:
            return img.size
    except Exception:
        return 0, 0


def optimize_image(data, source_name):
    with Image.open(BytesIO(data)) as img:
        img.load()
        width, height = img.size
        scale = min(1, MAX_IMAGE_EDGE / max(width, height))
        if scale < 1:
            img = img.resize((round(width * scale), round(height * scale)), Image.Resampling.LANCZOS)
        out = BytesIO()
        has_alpha = img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info)
        if has_alpha:
            img.save(out, format="PNG", optimize=True)
            return out.getvalue(), ".png"
        if img.mode != "RGB":
            img = img.convert("RGB")
        img.save(out, format="JPEG", quality=JPEG_QUALITY, optimize=True, progressive=True)
        return out.getvalue(), ".jpg"


def extract_images_for_sheet(zf, drawing_path):
    if drawing_path is None:
        return {}
    rel_path = posixpath.join(posixpath.dirname(drawing_path), "_rels", posixpath.basename(drawing_path) + ".rels")
    drawing_rels = rels_map(zf, rel_path)
    root = ET.fromstring(zf.read(drawing_path))
    images = {}
    for anchor in list(root):
        if not anchor.tag.endswith("Anchor"):
            continue
        from_node = anchor.find("xdr:from", NS)
        blip = anchor.find(".//a:blip", NS)
        if from_node is None or blip is None:
            continue
        row = int(from_node.find("xdr:row", NS).text)
        col = int(from_node.find("xdr:col", NS).text)
        rid = blip.attrib.get(f"{{{NS['officeRel']}}}embed")
        target = drawing_rels.get(rid)
        if not target:
            continue
        media_path = posixpath.normpath(posixpath.join(posixpath.dirname(drawing_path), target))
        images.setdefault((row, col), []).append(media_path)
    return images


def select_best_image(image_paths, media_cache):
    best = None
    best_score = -1
    for path in image_paths:
        data = media_cache[path]
        width, height = image_size(data)
        score = width * height
        if score > best_score:
            best_score = score
            best = path
    return best


def main():
    if len(sys.argv) != 3:
        raise SystemExit("Usage: rebuild_highres_site.py source.xlsx repo_dir")
    workbook_path = Path(sys.argv[1]).expanduser().resolve()
    repo_dir = Path(sys.argv[2]).resolve()
    index_path = repo_dir / "index.html"
    image_dir = repo_dir / "data" / "images"
    image_dir.mkdir(parents=True, exist_ok=True)
    for old in image_dir.glob("*"):
        if old.is_file():
            old.unlink()

    old_html = index_path.read_text(encoding="utf-8")
    old_products = json.loads(re.search(r"const products = ([\s\S]*?);\nlet currentKeyword", old_html).group(1))
    old_by_key = {(p.get("sheet"), p.get("code")): p for p in old_products}

    image_map = {}
    image_exports = {}
    missing_package_images = 0

    with zipfile.ZipFile(workbook_path) as zf:
        shared_strings = read_shared_strings(zf)
        sheets = parse_workbook_sheets(zf)
        media_cache = {name: zf.read(name) for name in zf.namelist() if name.startswith("xl/media/") and not name.endswith("/")}

        for sheet in sheets:
            if sheet["state"] != "visible" or sheet["name"] not in VISIBLE_SHEETS:
                continue
            sheet_root = ET.fromstring(zf.read(sheet["path"]))
            drawing_path = sheet_drawing_path(zf, sheet["path"], sheet_root)
            anchored_images = extract_images_for_sheet(zf, drawing_path)

            rows = {}
            for row_node in sheet_root.findall("main:sheetData/main:row", NS):
                row_index = int(row_node.attrib["r"]) - 1
                row_values = {}
                for cell in row_node.findall("main:c", NS):
                    pos = cell_ref_to_pos(cell.attrib.get("r", ""))
                    if pos is None:
                        continue
                    _, col_index = pos
                    row_values[col_index] = read_cell_value(cell, shared_strings)
                rows[row_index] = row_values

            header_row = None
            headers = {}
            for row_index in sorted(rows):
                inverse = {value.strip(): col for col, value in rows[row_index].items() if value}
                if "款式编码" in inverse and "主包装图片" in inverse:
                    header_row = row_index
                    headers = inverse
                    break
                if "款式编码" in inverse:
                    next_row = rows.get(row_index + 1, {})
                    merged = dict(rows[row_index])
                    merged.update({col: value for col, value in next_row.items() if value})
                    merged_inverse = {value.strip(): col for col, value in merged.items() if value}
                    if "主包装图片" in merged_inverse:
                        header_row = row_index + 1
                        headers = merged_inverse
                        break
            if header_row is None:
                continue

            style_col = headers.get("款式图")
            package_col = headers.get("主包装图片")

            for row_index in sorted(rows):
                if row_index <= header_row:
                    continue
                row = rows[row_index]
                code = str(row.get(headers["款式编码"], "")).strip()
                if not code:
                    continue

                product_images = {}

                for role, col in (("style", style_col), ("package", package_col)):
                    if col is None:
                        continue
                    candidates = anchored_images.get((row_index, col), [])
                    if not candidates:
                        continue
                    media_path = select_best_image(candidates, media_cache)
                    if not media_path:
                        continue
                    data = media_cache[media_path]
                    digest_key = (sheet["name"], code, role, media_path)
                    exported_name = image_exports.get(digest_key)
                    if exported_name is None:
                        optimized, ext = optimize_image(data, media_path)
                        safe_sheet = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff_-]+", "_", sheet["name"])
                        exported_name = f"{safe_sheet}_{code}_{role}{ext}"
                        (image_dir / exported_name).write_bytes(optimized)
                        image_exports[digest_key] = exported_name
                    product_images[f"{role}_image"] = f"data/images/{exported_name}"

                image_map[(sheet["name"], code)] = product_images
                if not product_images.get("package_image"):
                    missing_package_images += 1

    products = []
    for old in old_products:
        product = dict(old)
        mapped = image_map.get((product.get("sheet"), product.get("code")), {})
        if mapped.get("package_image"):
            product["package_image"] = mapped["package_image"]
        else:
            product["package_image"] = ""
        if mapped.get("style_image"):
            product["style_image"] = mapped["style_image"]
        products.append(product)

    css_patch = """
.image-wrap{position:relative;background:#f0f0f0;display:flex;align-items:center;justify-content:center;overflow:hidden}
.product-image{width:100%;height:200px;object-fit:contain;background:#f0f0f0}
.detail-images{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px}
.detail-images img{width:100%;max-width:100%;max-height:70vh;object-fit:contain;border-radius:8px;background:#f8f8f8;margin:0}
"""
    html = old_html
    html = re.sub(r"\.product-image\{[^}]+\}", ".product-image{width:100%;height:200px;object-fit:contain;background:#f0f0f0}", html)
    html = re.sub(r"\.detail-images img\{[^}]+\}", ".detail-images img{width:100%;max-width:100%;max-height:70vh;object-fit:contain;border-radius:8px;background:#f8f8f8;margin:0}", html)
    if ".image-wrap" not in html:
        html = html.replace("</style>", css_patch + "\n</style>")

    products_json = json.dumps(products, ensure_ascii=False)
    html = re.sub(
        r"const products = [\s\S]*?;\nlet currentKeyword",
        lambda _match: "const products = " + products_json + ";\nlet currentKeyword",
        html,
    )
    html = html.replace(
        "var img = p.package_image && p.package_image.startsWith('data:') ? '<img class=\"product-image\" src=\"' + p.package_image + '\">' : '<div class=\"product-image\" style=\"display:flex;align-items:center;justify-content:center;color:#999\">暂无图片</div>';",
        "var imageSrc = p.package_image || p.style_image || ''; var img = imageSrc ? '<div class=\"image-wrap\"><img class=\"product-image\" src=\"' + imageSrc + '\" loading=\"lazy\"></div>' : '<div class=\"product-image\" style=\"display:flex;align-items:center;justify-content:center;color:#999\">暂无图片</div>';",
    )
    html = html.replace(
        "if (p.package_image && p.package_image.startsWith('data:')) html += '<div class=\"detail-section\"><h3>商品图片</h3><div class=\"detail-images\"><img src=\"' + p.package_image + '\"></div></div>';",
        "var detailImgs = [p.package_image, p.style_image].filter(function(src, i, arr){return src && arr.indexOf(src) === i}); if (detailImgs.length) html += '<div class=\"detail-section\"><h3>商品图片</h3><div class=\"detail-images\">' + detailImgs.map(function(src){return '<a href=\"' + src + '\" target=\"_blank\"><img src=\"' + src + '\" loading=\"lazy\"></a>'}).join('') + '</div></div>';",
    )
    html = html.replace(
        'document.getElementById("stats").innerHTML = "共 <span>463</span> 个，当前 <span>" + list.length + "</span> 个";',
        'document.getElementById("stats").innerHTML = "共 <span>" + products.length + "</span> 个，当前 <span>" + list.length + "</span> 个";',
    )
    index_path.write_text(html, encoding="utf-8")

    print(json.dumps({
        "products": len(products),
        "images_written": len(list(image_dir.glob("*"))),
        "missing_package_images": missing_package_images,
        "output": str(index_path),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
