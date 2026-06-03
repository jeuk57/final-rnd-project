"""입력 문서에서 텍스트를 추출해 state['extracted_text']에 저장.

지원:
- PDF: document_parsing.extract_text_from_pdf
- DOCX: document_parsing.parse_docx_to_blocks
- 이미 파싱된 JSON: *_parsing.json

주의:
- 이 프로젝트의 기존 파서(document_parsing.py)를 그대로 활용한다.
"""

from __future__ import annotations

import json
import os
from typing import Any, Dict, List

from utils.document_parsing import extract_text_from_pdf, parse_docx_to_blocks


def _flatten_pdf_pages(pages: List[Dict[str, Any]]) -> str:
    # extract_text_from_pdf는 페이지별 texts 리스트를 반환
    out: List[str] = []
    for p in pages:
        texts = p.get("texts") or []
        for t in texts:
            if t:
                out.append(str(t))
        out.append("\n")
    return "\n".join(out)


def _flatten_docx_blocks(parsed: Dict[str, Any]) -> str:
    blocks = parsed.get("blocks") or []
    out: List[str] = []
    for b in blocks:
        if not b:
            continue
        t = b.get("type")
        if t in {"paragraph", "textbox", "table", "text"}:
            if "text" in b and b["text"]:
                out.append(str(b["text"]))
        elif t == "image":
            out.append("[IMAGE: 문서에 이미지/도표 포함]")
    return "\n".join(out)

def extract_text(state: Dict[str, Any]) -> Dict[str, Any]:
    src = state.get("source_path")
    if not src:
        raise RuntimeError("source_path가 없습니다. 입력 파일 경로가 필요합니다.")

    if not os.path.exists(src):
        raise FileNotFoundError(src)

    ext = os.path.splitext(src)[1].lower()
    extracted_text = ""

    if ext == ".json":
        with open(src, "r", encoding="utf-8") as f:
            data = json.load(f)
        # 개발용 재실행 흐름: 이미 저장된 파싱 JSON의 문자열 입력 재사용
        # pdf parsing json은 list 형태, docx parsing json은 dict 형태
        if isinstance(data, list):
            extracted_text = _flatten_pdf_pages(data)
        elif isinstance(data, dict):
            extracted_text = _flatten_docx_blocks(data)
        else:
            extracted_text = str(data)
    elif ext == ".pdf":
        pages = extract_text_from_pdf(src)
        extracted_text = _flatten_pdf_pages(pages)
    elif ext == ".docx":
        out_dir = state.get("parsing_out_dir") or os.path.join(os.getcwd(), "parsing")
        parsed = parse_docx_to_blocks(src, out_dir)
        extracted_text = _flatten_docx_blocks(parsed)
    else:
        raise RuntimeError(f"지원하지 않는 입력 형식입니다: {ext} (pdf/docx/json 지원)")

    state["extracted_text"] = extracted_text
    print("[DEBUG] state keys:", list(state.keys()))
    print("[DEBUG] source_path:", state.get("source_path"))
    print("[DEBUG] extracted_text length:", len(extracted_text) if extracted_text else 0)

    return state
