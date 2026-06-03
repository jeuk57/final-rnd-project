import os
import re
import zipfile
import pdfplumber
from operator import itemgetter
from lxml import etree
from typing import Any, Dict, List

# ==========================================================
# 1. WORD 파싱 로직
# ==========================================================
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "v": "urn:schemas-microsoft-com:vml",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def read_xml(z: zipfile.ZipFile, path: str) -> etree._Element:
    return etree.fromstring(z.read(path))


def parse_document_rels(z: zipfile.ZipFile) -> Dict[str, str]:
    rels_path = "word/_rels/document.xml.rels"
    if rels_path not in z.namelist():
        return {}
    root = read_xml(z, rels_path)
    rid_to_target = {}
    for rel in root.findall("rel:Relationship", namespaces=NS):
        rid = rel.get("Id")
        target = rel.get("Target")
        if rid and target:
            rid_to_target[rid] = target
    return rid_to_target


def get_text_from_runs(p_elm: etree._Element) -> str:
    texts = [t.text for t in p_elm.findall(".//w:t", namespaces=NS) if t.text]
    return "".join(texts).strip()


def extract_image_rids_from_paragraph(p_elm: etree._Element) -> List[str]:
    rids = []
    for blip in p_elm.findall(".//a:blip", namespaces=NS):
        rid = blip.get(f"{{{NS['r']}}}embed")
        if rid:
            rids.append(rid)

    seen = set()
    uniq = []
    for rid in rids:
        if rid not in seen:
            uniq.append(rid)
            seen.add(rid)
    return uniq


def save_image_by_rid(z, rid, rid_to_target, media_out_dir, index):
    target = rid_to_target.get(rid)
    if not target:
        return None

    zip_img_path = target if target.startswith("word/") else f"word/{target}"
    if zip_img_path not in z.namelist():
        return None

    img_bytes = z.read(zip_img_path)
    base_name = os.path.basename(zip_img_path)
    safe_base = re.sub(r"[^a-zA-Z0-9._-]+", "_", base_name)
    out_name = f"{index:04d}_{safe_base}"
    out_path = os.path.join(media_out_dir, out_name)

    with open(out_path, "wb") as f:
        f.write(img_bytes)

    return {
        "type": "image",
        "rid": rid,
        "path_in_docx": zip_img_path,
        "saved_as": out_path.replace("\\", "/"),
        "bytes": len(img_bytes),
    }


def extract_textboxes_from_paragraph(p_elm: etree._Element) -> List[str]:
    results = []
    vml = [t.text for t in p_elm.findall(".//w:pict//v:textbox//w:t", namespaces=NS) if t.text]
    if vml:
        results.append("".join(vml).strip())

    draw = [t.text for t in p_elm.findall(".//w:drawing//a:t", namespaces=NS) if t.text]
    if draw:
        results.append("".join(draw).strip())

    wps = [t.text for t in p_elm.findall(".//wps:txbx//w:t", namespaces=NS) if t.text]
    if wps:
        results.append("".join(wps).strip())

    uniq = []
    seen = set()
    for s in results:
        s2 = s.strip()
        if s2 and s2 not in seen:
            uniq.append(s2)
            seen.add(s2)
    return uniq


def parse_table(tbl_elm: etree._Element) -> Dict[str, Any]:
    rows = []
    for tr in tbl_elm.findall(".//w:tr", namespaces=NS):
        row_cells = [
            "".join([t.text for t in tc.findall(".//w:t", namespaces=NS) if t.text]).strip()
            for tc in tr.findall(".//w:tc", namespaces=NS)
        ]
        rows.append(row_cells)
    return {"type": "table", "rows": rows}


def parse_docx_to_blocks(docx_path: str, out_dir: str) -> Dict[str, Any]:
    # 이미지 저장 폴더는 parsing/media_파일명 형식으로 분리
    media_out_dir = os.path.join(out_dir, "media_" + os.path.basename(docx_path))
    os.makedirs(media_out_dir, exist_ok=True)

    with zipfile.ZipFile(docx_path) as z:
        doc_root = read_xml(z, "word/document.xml")
        rid_to_target = parse_document_rels(z)
        body = doc_root.find(".//w:body", namespaces=NS)
        blocks = []
        img_counter = 0

        for child in body:
            tag = etree.QName(child).localname
            if tag == "p":
                text = get_text_from_runs(child)
                if text:
                    blocks.append({"type": "paragraph", "text": text})
                for tb in extract_textboxes_from_paragraph(child):
                    blocks.append({"type": "textbox", "text": tb})
                for rid in extract_image_rids_from_paragraph(child):
                    img_counter += 1
                    img_block = save_image_by_rid(z, rid, rid_to_target, media_out_dir, img_counter)
                    blocks.append(img_block if img_block else {"type": "image_ref", "rid": rid, "note": "missing"})
            elif tag == "tbl":
                blocks.append(parse_table(child))

    return {"source": os.path.basename(docx_path), "blocks": blocks}


# ==========================================================
# 2. PDF 파싱 로직
# ==========================================================
def filter_overlapping_tables(tables):
    if not tables:
        return []

    indices_to_remove = set()
    for i, outer in enumerate(tables):
        for j, inner in enumerate(tables):
            if i == j:
                continue
            if (
                outer.bbox[0] <= inner.bbox[0] + 1
                and outer.bbox[1] <= inner.bbox[1] + 1
                and outer.bbox[2] >= inner.bbox[2] - 1
                and outer.bbox[3] >= inner.bbox[3] - 1
            ):
                indices_to_remove.add(i)
                break
    return [t for i, t in enumerate(tables) if i not in indices_to_remove]


def is_inside_bbox(word, bboxes):
    w_cx = (word["x0"] + word["x1"]) / 2
    w_cy = (word["top"] + word["bottom"]) / 2
    for b in bboxes:
        if (b[0] <= w_cx <= b[2]) and (b[1] <= w_cy <= b[3]):
            return True
    return False


def table_to_markdown(table_data):
    if not table_data:
        return ""
    return "\n".join(
        [
            "| " + " | ".join([str(cell).replace("\n", " ").strip() if cell else "" for cell in row]) + " |"
            for row in table_data
        ]
    )


def extract_text_from_pdf(pdf_path):
    all_pages_data = []
    doc_id = os.path.basename(pdf_path)

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            raw_tables = page.find_tables()
            tables = filter_overlapping_tables(raw_tables)
            table_bboxes = [t.bbox for t in tables]
            page_contents = []

            for table in tables:
                extracted_data = table.extract()
                if not extracted_data:
                    continue

                if len(extracted_data) == 1 and len(extracted_data[0]) == 1:
                    page_contents.append(
                        {
                            "type": "text",
                            "top": table.bbox[1],
                            "text": str(extracted_data[0][0]).strip().replace("\n", " "),
                        }
                    )
                else:
                    md_table = table_to_markdown(extracted_data)
                    if md_table:
                        page_contents.append(
                            {
                                "type": "table",
                                "top": table.bbox[1],
                                "text": f"\n[TABLE START]\n{md_table}\n[TABLE END]",
                            }
                        )

            for img in page.images:
                if img["height"] > 10 and img["width"] > 10:
                    page_contents.append(
                        {
                            "type": "image",
                            "top": img["top"],
                            "text": "\n[IMAGE: 그림/도표/이미지 포함]\n",
                        }
                    )

            words = page.extract_words()
            words_outside_tables = [w for w in words if not is_inside_bbox(w, table_bboxes)]
            if words_outside_tables:
                lines = []
                current_line = [words_outside_tables[0]]
                for i in range(1, len(words_outside_tables)):
                    if abs(words_outside_tables[i]["top"] - words_outside_tables[i - 1]["top"]) < 5:
                        current_line.append(words_outside_tables[i])
                    else:
                        lines.append(current_line)
                        current_line = [words_outside_tables[i]]
                lines.append(current_line)

                for line in lines:
                    merged_text = " ".join([w["text"] for w in line]).strip()
                    if merged_text:
                        page_contents.append({"type": "text", "top": line[0]["top"], "text": merged_text})

            page_contents.sort(key=itemgetter("top"))
            all_pages_data.append(
                {
                    "doc_id": doc_id,
                    "page_index": page_idx,
                    "texts": [item["text"] for item in page_contents],
                }
            )

    return all_pages_data


# ==========================================================
# 3. 통합 실행 로직
# ==========================================================
if __name__ == "__main__":
    # 경로 설정
    input_base_dir = "data/input"
    output_base_dir = "parsing"

    os.makedirs(input_base_dir, exist_ok=True)
    os.makedirs(output_base_dir, exist_ok=True)

    # input 폴더의 모든 파일 대상
    for filename in os.listdir(input_base_dir):
        input_path = os.path.join(input_base_dir, filename)
        if not os.path.isfile(input_path):
            continue

        name_only, ext = os.path.splitext(filename)
        ext = ext.lower()

        # 출력 파일명 규칙: 원래 파일명_parsing.json
        output_filename = f"{name_only}_parsing.json"
        output_path = os.path.join(output_base_dir, output_filename)

        print(f"작업 시작: {filename}")

        try:
            if ext == ".docx":
                result = parse_docx_to_blocks(input_path, output_base_dir)
            elif ext == ".pdf":
                result = extract_text_from_pdf(input_path)
            else:
                print(f"건너뜀 (지원하지 않는 형식): {filename}")
                continue

            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(result, f, ensure_ascii=False, indent=2)
            print(f"완료: {output_path}")

        except Exception as e:
            print(f"오류 발생 ({filename}): {e}")
