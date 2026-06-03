import json
import os
import zipfile
from operator import itemgetter
from typing import Any, Dict, List

import pdfplumber
from lxml import etree

# =========================
# Word(OpenXML) 네임스페이스
# =========================
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "v": "urn:schemas-microsoft-com:vml",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


class UniversalParser:
    def __init__(self, output_dir: str = "output"):
        self.output_dir = output_dir
        os.makedirs(self.output_dir, exist_ok=True)

    # ---------------------------------------------------------
    # PDF 파싱 로직
    # ---------------------------------------------------------
    def _filter_overlapping_tables(self, tables):
        if not tables:
            return []

        indices_to_remove = set()
        for i, outer in enumerate(tables):
            outer_bbox = outer.bbox
            for j, inner in enumerate(tables):
                if i == j:
                    continue
                inner_bbox = inner.bbox
                if (
                    outer_bbox[0] <= inner_bbox[0] + 1
                    and outer_bbox[1] <= inner_bbox[1] + 1
                    and outer_bbox[2] >= inner_bbox[2] - 1
                    and outer_bbox[3] >= inner_bbox[3] - 1
                ):
                    indices_to_remove.add(i)
                    break

        return [t for i, t in enumerate(tables) if i not in indices_to_remove]

    def _is_inside_bbox(self, word, bboxes):
        w_center_x = (word["x0"] + word["x1"]) / 2
        w_center_y = (word["top"] + word["bottom"]) / 2
        for bbox in bboxes:
            if bbox[0] <= w_center_x <= bbox[2] and bbox[1] <= w_center_y <= bbox[3]:
                return True
        return False

    def _table_to_markdown(self, table_data):
        if not table_data:
            return ""
        lines = []
        for row in table_data:
            cleaned = [str(cell).replace("\n", " ").strip() if cell else "" for cell in row]
            lines.append("| " + " | ".join(cleaned) + " |")
        return "\n".join(lines)

    def parse_pdf(self, pdf_path: str) -> List[Dict]:
        doc_data = []
        doc_id = os.path.basename(pdf_path)

        with pdfplumber.open(pdf_path) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                raw_tables = page.find_tables()
                tables = self._filter_overlapping_tables(raw_tables)
                table_bboxes = [t.bbox for t in tables]
                page_contents = []

                # 표 추출
                for table in tables:
                    extracted = table.extract()
                    if not extracted:
                        continue

                    if len(extracted) == 1 and len(extracted[0]) == 1:
                        text = str(extracted[0][0]).strip().replace("\n", " ")
                        if text:
                            page_contents.append({"type": "text", "top": table.bbox[1], "text": text})
                    else:
                        md = self._table_to_markdown(extracted)
                        if md:
                            page_contents.append(
                                {
                                    "type": "table",
                                    "top": table.bbox[1],
                                    "text": f"[TABLE]\n{md}",
                                }
                            )

                # 이미지 영역 표시
                for img in page.images:
                    if img.get("height", 0) > 10 and img.get("width", 0) > 10:
                        page_contents.append(
                            {
                                "type": "image",
                                "top": img.get("top", 0),
                                "text": "[IMAGE]",
                            }
                        )

                # 표 영역 밖 텍스트 추출
                words = page.extract_words()
                words = [w for w in words if not self._is_inside_bbox(w, table_bboxes)]

                if words:
                    words.sort(key=itemgetter("top", "x0"))
                    lines = []
                    curr = [words[0]]

                    for w in words[1:]:
                        if abs(w["top"] - curr[-1]["top"]) < 5:
                            curr.append(w)
                        else:
                            lines.append(curr)
                            curr = [w]
                    lines.append(curr)

                    for line in lines:
                        merged = " ".join(w["text"] for w in line).strip()
                        if merged:
                            page_contents.append({"type": "text", "top": line[0]["top"], "text": merged})

                page_contents.sort(key=itemgetter("top"))

                doc_data.append(
                    {
                        "doc_id": doc_id,
                        "page_index": page_idx,
                        "contents": [c["text"] for c in page_contents],
                    }
                )

        return doc_data

    # ---------------------------------------------------------
    # DOCX 파싱 로직
    # ---------------------------------------------------------
    def _read_xml(self, z, path):
        return etree.fromstring(z.read(path))

    def parse_docx(self, docx_path: str) -> Dict:
        with zipfile.ZipFile(docx_path) as z:
            if "word/document.xml" not in z.namelist():
                return {"error": "Invalid docx"}

            root = self._read_xml(z, "word/document.xml")
            body = root.find(".//w:body", namespaces=NS)

            blocks = []

            for child in body:
                tag = etree.QName(child).localname

                if tag == "p":
                    text = "".join(
                        t.text for t in child.findall(".//w:t", namespaces=NS) if t.text
                    ).strip()
                    if text:
                        blocks.append({"type": "paragraph", "text": text})

                elif tag == "tbl":
                    rows = []
                    for tr in child.findall(".//w:tr", namespaces=NS):
                        row = [
                            "".join(t.text for t in tc.findall(".//w:t", namespaces=NS) if t.text).strip()
                            for tc in tr.findall(".//w:tc", namespaces=NS)
                        ]
                        rows.append(row)
                    blocks.append({"type": "table", "rows": rows})

            return {
                "source": os.path.basename(docx_path),
                "blocks": blocks,
            }


# =========================================================
# FastAPI에서 직접 쓰는 파싱 진입점
# =========================================================
def parse_file_to_json(file_path: str) -> Any:
    """
    파일 경로를 받아 파싱된 JSON 객체를 반환한다.
    """
    parser = UniversalParser()
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".pdf":
        return {
            "file_type": "pdf",
            "pages": parser.parse_pdf(file_path),
        }

    if ext == ".docx":
        return {
            "file_type": "docx",
            "content": parser.parse_docx(file_path),
        }

    return {
        "error": f"Unsupported extension: {ext}",
    }
