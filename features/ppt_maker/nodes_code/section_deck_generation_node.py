from __future__ import annotations

import json
import re
from typing import Any, Dict, List

from google import genai
from google.genai import types

from .llm_utils import generate_content_with_retry, get_gemini_client


# -----------------------------
# 프롬프트 설정
# -----------------------------
FORBIDDEN_ENDINGS = ("다", "니다", "합니다", "됩니다")
FORMAL_LINE_RE = re.compile(
    r"(합니다|입니다|됩니다|있습니다|가능합니다|예상됩니다|필요합니다|수행합니다|한다|된다|있다)\s*[.!?]?$"
)


def _contains_formal_line(text: Any) -> bool:
    t = str(text or "").strip()
    if not t:
        return False
    for line in t.splitlines():
        if FORMAL_LINE_RE.search(line.strip()):
            return True
    return False


# Gemini 섹션별 슬라이드 생성 프롬프트 구성 함수
def _build_prompt() -> str:
    return """
역할: 너는 국가 R&D 선정평가 발표자료(PPT)를 작성하는 실무 PM/총괄기획자다.

목표:
'AI가 만든 티'가 아니라 실제 선정평가장에서 쓰는 자료처럼,
보편적·사실적·증빙가능한 근거 중심으로
논리 완결성을 갖춘 발표자료를 작성한다.

제약:
- 지금 입력에는 특정 '섹션'의 원문만 들어있다.
- 반드시 그 섹션에 해당하는 슬라이드만 생성한다.
- 근거/수치가 없으면 추정하지 말고 '미기재'로 둔다.
- 출력은 아래 포맷을 100% 지켜라. (추가 텍스트 금지)

중요 원칙:
슬라이드 수를 최소화하려 하지 말고,
발표자가 실제로 설명할 수 있는 수준까지 충분히 분리하여 작성한다.

--------------------------------

출력 포맷:

DECK_TITLE: <발표자료 전체 제목 1줄>

DECK_TITLE 작성 규칙:
- DECK_TITLE은 현재 섹션 제목이 아니라 전체 PPT 제목이다.
- 현재 입력 원문에서 전체 과제명 또는 전체 발표자료 제목을 명확히 확인할 수 있는 경우에만 작성한다.
- 확인할 수 없으면 추정하지 말고 '(과제명 미기재)'로 작성한다.
- 현재 섹션 내용만 보고 새로운 제목을 임의 생성하지 않는다.

SLIDE
SECTION: <섹션명>
TITLE: <슬라이드 제목 1줄>
KEY_MESSAGE: <핵심 메시지 1줄>

BULLETS:
- <불릿 1>
- <불릿 2>
- <불릿 3 이상 작성>

EVIDENCE:
- type: <출처/수치/근거>
  text: <텍스트>

IMAGE_NEEDED: <true/false>
IMAGE_TYPE: <diagram/chart/table/none 중 하나>
IMAGE_BRIEF_KO: <(사진/일러스트 금지) 벡터 인포그래픽/도형/차트 지시문 (없으면 빈 문자열)>

TABLE_MD: <마크다운 표(여러 줄 가능). 없으면 빈 문자열>
DIAGRAM_SPEC_KO: <도형 기반 도식 스펙(여러 줄 가능). 없으면 빈 문자열>
CHART_SPEC_KO: <차트 스펙(여러 줄 가능). 없으면 빈 문자열>

ENDSLIDE

--------------------------------

절대 금지:
- CHAPTER / PART / SECTION 같은 구분용 단독 슬라이드 생성 금지
- 표지 전용 슬라이드 생성 금지
- 사진 / 실사 / 캐릭터 / 3D / AI그림 금지
- (중요) IMAGE_NEEDED는 항상 false. 이미지 파일 생성 지시 금지. (텍스트 없는 아이콘/도형은 PPT에서 수동 삽입 가능)
- 허용: 도형 기반 인포그래픽, 차트, 표

--------------------------------

시각요소 규칙:
- 시각요소(TABLE_MD / DIAGRAM_SPEC_KO / CHART_SPEC_KO)는 선택 사항이다. 없어도 된다.
- 넣더라도 '이미지 생성'이 아니라, 발표자가 직접 도형/표를 그릴 수 있을 만큼의 텍스트 지시서만 작성한다.

추가 제약(최우선):
- 출력의 모든 문장/표현은 한국어로 작성한다.
- 영어 문장/영어 소제목/영어 불릿 금지.
- 단, 고유명사/약어(API, GPU 등)만 예외적으로 허용.
- 가능한 한 '한글 용어(괄호에 약어)' 형태로 쓴다. 예: 그래픽처리장치(GPU)
- 문장 종결은 발표 메모형으로 작성한다. (예: ~확보, ~예정, ~검토, ~적용)
- '~입니다/~합니다/~하였다' 같은 서술형 종결문은 사용하지 않는다.
- 불릿은 '명사형/행동형 메모체'로 작성한다. (예: 데이터 확보, 일정 검토, 적용 예정)

[필수 규칙 - 발표자료 문체]
- 문장 형태로 작성하지 않는다.
- 모든 항목은 명사구 또는 키워드 형태로 작성한다.
- 문장 종결어미 사용 금지 (~다, ~니다, ~합니다, ~됩니다 포함)
- 발표 슬라이드용 불릿 형태로 작성한다.
- 최소 3개의 불릿이 없는 경우 슬라이드를 생성하지 않는다.
- 내용이 부족하면 슬라이드를 만들지 않는다.
- 목차 슬라이드에서는 제목만 출력한다. 설명 문장은 출력하지 않는다.
TITLE 규칙:
- TITLE은 짧게 작성한다.
- TITLE은 최대 2줄을 넘기지 않는다.
- TITLE에 긴 문장을 만들지 않는다.
- TITLE에 영어 문장을 쓰지 않는다.
- TITLE은 단순한 한국어 제목 형태로 작성한다.
KEY_MESSAGE 추가 제약:
- KEY_MESSAGE는 문장 형태로 작성하지 않는다.
- KEY_MESSAGE는 정확히 3개의 키워드/명사구로 작성한다.
- KEY_MESSAGE 형식: 키워드1, 키워드2, 키워드3
- TITLE/KEY_MESSAGE/BULLETS/EVIDENCE의 서술형 종결 금지:
  합니다, 입니다, 됩니다, 있습니다, 가능합니다, 예상됩니다, 필요합니다, 수행합니다, 한다, 된다, 있다
""".strip()


# -----------------------------
# 파싱 보조 함수
# -----------------------------
def _parse_deck_title(raw: str) -> str:
    m = re.search(r"(?m)^DECK_TITLE:\s*(.+)$", raw)
    return (m.group(1).strip() if m else "").strip()


def _iter_slide_blocks(raw: str) -> List[str]:
    blocks: List[str] = []
    pattern = re.compile(r"(?s)\bSLIDE\b(.*?)\bENDSLIDE\b")
    for m in pattern.finditer(raw):
        blocks.append(m.group(1).strip())
    return blocks


def _grab_field(block: str, field: str) -> str:
    m = re.search(rf"(?m)^{re.escape(field)}\s*:\s*(.+)$", block)
    return (m.group(1).strip() if m else "").strip()


def _grab_multiline_field(block: str, field: str) -> str:
    m = re.search(
        rf"(?s)^{re.escape(field)}\s*:\s*(.*?)(?=\n[A-Z_]+\s*:|\Z)",
        block.strip(),
        flags=re.MULTILINE,
    )
    return (m.group(1).strip() if m else "").strip()


def _parse_bullets(block: str) -> List[str]:
    m = re.search(r"(?s)\bBULLETS\s*:\s*(.*?)(?:\n[A-Z_]+\s*:|\Z)", block)
    if not m:
        return []
    body = m.group(1)
    bullets: List[str] = []
    for line in body.splitlines():
        line = line.strip()
        if line.startswith("-"):
            t = line[1:].strip()
            if t:
                bullets.append(t)
    return bullets


def _parse_evidence(block: str) -> List[Dict[str, str]]:
    m = re.search(r"(?s)\bEVIDENCE\s*:\s*(.*?)(?:\n[A-Z_]+\s*:|\Z)", block)
    if not m:
        return []
    body = m.group(1)
    items: List[Dict[str, str]] = []

    cur: Dict[str, str] = {}
    for line in body.splitlines():
        line = line.rstrip()
        if not line.strip():
            continue

        if line.lstrip().startswith("- type:"):
            if cur:
                items.append(cur)
            cur = {"type": line.split(":", 1)[1].strip(), "text": ""}
            continue

        if line.strip().startswith("text:"):
            cur["text"] = line.split(":", 1)[1].strip()
            continue

        if line.lstrip().startswith("-"):
            t = line.lstrip()[1:].strip()
            if t:
                items.append({"type": "근거", "text": t})
            cur = {}
            continue

    if cur:
        items.append(cur)

    cleaned: List[Dict[str, str]] = []
    for it in items:
        t = (it.get("text") or "").strip()
        if t:
            cleaned.append({"type": (it.get("type") or "근거").strip(), "text": t})
    return cleaned


def _parse_bool(s: str) -> bool:
    return (s or "").strip().lower() in {"true", "1", "yes", "y"}


def _to_phrase(text: Any) -> str:
    t = re.sub(r"\s+", " ", str(text or "")).strip()
    if not t:
        return ""
    t = re.sub(r"[.!?]+$", "", t).strip()
    for end in FORBIDDEN_ENDINGS:
        if t.endswith(end):
            t = t[: -len(end)].strip()
            break
    return t


def _keyword_tokens(text: Any) -> List[str]:
    src = str(text or "").replace("\n", ",")
    parts = re.split(r"[,/|·;]", src)
    out: List[str] = []
    for p in parts:
        k = _to_phrase(p)
        k = re.sub(r"^\d+[\.\)]\s*", "", k).strip()
        k = re.sub(r"\s{2,}", " ", k).strip()
        if not k:
            continue
        if _contains_formal_line(k):
            continue
        out.append(k)
    return out


def _format_key_message(key_message: Any, title: Any, bullets: List[Any]) -> str:
    candidates: List[str] = []
    candidates.extend(_keyword_tokens(key_message))
    for b in (bullets or [])[:6]:
        candidates.extend(_keyword_tokens(b))
    candidates.extend(_keyword_tokens(title))

    uniq: List[str] = []
    seen = set()
    for c in candidates:
        key = re.sub(r"\s+", "", c.lower())
        if not key or key in seen:
            continue
        seen.add(key)
        uniq.append(c)
        if len(uniq) >= 3:
            break

    while len(uniq) < 3:
        fallback = ["핵심 과제", "추진 방향", "운영 계획"][len(uniq)]
        uniq.append(fallback)
    return ", ".join(uniq[:3])


def _slide_has_formal_lines(slide: Dict[str, Any]) -> bool:
    if _contains_formal_line(slide.get("slide_title")):
        return True
    if _contains_formal_line(slide.get("key_message")):
        return True
    for b in (slide.get("bullets") or []):
        if _contains_formal_line(b):
            return True
    for ev in (slide.get("evidence") or []):
        if isinstance(ev, dict) and _contains_formal_line(ev.get("text")):
            return True
    return False


# 문장형 슬라이드 표현의 Gemini 재작성 함수
def _rewrite_formal_lines_with_gemini(
    client: genai.Client,
    model: str,
    slide: Dict[str, Any],
) -> Dict[str, Any]:
    prompt = (
        "슬라이드 텍스트를 발표자료용 키워드 형태로만 재작성한다.\n"
        "규칙:\n"
        "- 문장형 종결 금지.\n"
        "- 합니다/입니다/됩니다/있습니다 같은 서술형 종결 금지.\n"
        "- KEY_MESSAGE는 정확히 3개의 키워드 구문으로 작성.\n"
        "- BULLETS는 짧은 명사구로 작성.\n"
        "- EVIDENCE text는 짧은 명사구로 작성.\n"
        "title, key_message_keywords, bullets, evidence 키를 가진 JSON만 반환.\n"
    )
    payload = {
        "title": str(slide.get("slide_title") or ""),
        "key_message": str(slide.get("key_message") or ""),
        "bullets": slide.get("bullets") or [],
        "evidence": slide.get("evidence") or [],
    }
    resp = generate_content_with_retry(
        client,
        model=model or "gemini-2.5-flash",
        contents=[prompt, json.dumps(payload, ensure_ascii=False)],
        config=types.GenerateContentConfig(
            temperature=0.2,
            max_output_tokens=1024,
            response_mime_type="application/json",
        ),
        max_retries=1,
    )
    raw = (getattr(resp, "text", None) or "").strip()
    obj = json.loads(raw) if raw else {}
    if not isinstance(obj, dict):
        return slide

    out = dict(slide)
    out["slide_title"] = _to_phrase(obj.get("title") or out.get("slide_title") or "")
    km_list = obj.get("key_message_keywords") or []
    if isinstance(km_list, list):
        km_candidates = [_to_phrase(x) for x in km_list if _to_phrase(x)]
    else:
        km_candidates = _keyword_tokens(km_list)
    out["key_message"] = _format_key_message(", ".join(km_candidates), out.get("slide_title"), out.get("bullets") or [])

    bullets = obj.get("bullets") or out.get("bullets") or []
    out["bullets"] = [_to_phrase(x) for x in bullets if _to_phrase(x) and not _contains_formal_line(x)]

    ev_raw = obj.get("evidence") or out.get("evidence") or []
    ev_out: List[Dict[str, str]] = []
    if isinstance(ev_raw, list):
        for ev in ev_raw:
            if isinstance(ev, dict):
                t = _to_phrase(ev.get("text"))
                if t and not _contains_formal_line(t):
                    ev_out.append({"type": str(ev.get("type") or "근거").strip(), "text": t})
    out["evidence"] = ev_out
    return out


# Gemini 응답의 SLIDE/ENDSLIDE 블록 파싱 함수
def _parse_slides_from_text(raw: str, *, default_section: str, start_order: int) -> List[Dict[str, Any]]:
    slides: List[Dict[str, Any]] = []
    order = start_order

    for block in _iter_slide_blocks(raw):
        section = (_grab_field(block, "SECTION") or default_section).strip()
        title = _grab_field(block, "TITLE").strip()
        key_message = _grab_field(block, "KEY_MESSAGE").strip()

        image_needed = _parse_bool(_grab_field(block, "IMAGE_NEEDED"))
        image_type = (_grab_field(block, "IMAGE_TYPE") or "none").strip().lower()
        if image_type not in {"diagram", "chart", "table", "none"}:
            image_type = "none"

        image_brief_ko = _grab_multiline_field(block, "IMAGE_BRIEF_KO")

        # 이미지 생성 비활성화: Gemini 이미지 생성 대신 PPT 후처리/도형 기반 처리
        image_needed = False
        image_type = "none"
        image_brief_ko = ""

        table_md = _grab_multiline_field(block, "TABLE_MD")
        diagram_spec_ko = _grab_multiline_field(block, "DIAGRAM_SPEC_KO")
        chart_spec_ko = _grab_multiline_field(block, "CHART_SPEC_KO")

        bullets = [_to_phrase(b) for b in _parse_bullets(block) if _to_phrase(b)]
        evidence = _parse_evidence(block)

        # 챕터/파트/섹션 단독 슬라이드 제거
        upper_title = (title or "").upper()
        upper_section = (section or "").upper()
        if any(x in upper_title for x in ["CHAPTER", "PART", "SECTION"]) or any(
            x in upper_section for x in ["CHAPTER", "PART", "SECTION"]
        ):
            continue
        if len(bullets) < 3:
            continue

        slides.append(
            {
                "order": order,
                "section": section,
                "slide_title": _to_phrase(title) or title,
                "key_message": _format_key_message(key_message, title, bullets),
                "bullets": bullets,
                "evidence": evidence,
                "image_needed": False,
                "image_type": "none",
                "image_brief_ko": "",
                "TABLE_MD": table_md,
                "DIAGRAM_SPEC_KO": diagram_spec_ko,
                "CHART_SPEC_KO": chart_spec_ko,
            }
        )
        order += 1

    return slides


def _fallback_slide_from_raw(raw: str, *, default_section: str, order: int) -> List[Dict[str, Any]]:
    """
    Gemini 응답이 SLIDE/ENDSLIDE 형식을 벗어난 경우의 최소 슬라이드 복구 함수.
    섹션 전체 누락 방지 목적.
    """
    txt = (raw or "").strip()
    if not txt:
        return []

    title = default_section or "핵심 내용"
    key_message = ""
    bullets: List[str] = []

    for line in txt.splitlines():
        line = line.strip()
        if not line:
            continue
        if not key_message and len(line) <= 80 and not line.startswith("-"):
            key_message = line
            continue
        if line.startswith("-"):
            b = line[1:].strip()
            if b:
                bullets.append(b)
        elif len(line) <= 120:
            bullets.append(line)
        if len(bullets) >= 4:
            break

    if not key_message:
        key_message = title
    bullets = [_to_phrase(b) for b in bullets if _to_phrase(b)]
    key_message = _format_key_message(key_message, title, bullets)
    if len(bullets) < 3:
        return []

    return [
        {
            "order": order,
            "section": default_section,
            "slide_title": title,
            "key_message": key_message,
            "bullets": bullets[:4],
            "evidence": [],
            "image_needed": False,
            "image_type": "none",
            "image_brief_ko": "",
            "TABLE_MD": "",
            "DIAGRAM_SPEC_KO": "",
            "CHART_SPEC_KO": "",
        }
    ]


# 긴 섹션 원문의 LLM 입력 분할 함수
def _split_section_text_for_llm(text: str, *, max_chunk_chars: int, max_chunks: int) -> List[str]:
    txt = (text or "").strip()
    if not txt:
        return []
    if len(txt) <= max_chunk_chars:
        return [txt]

    paras = [p.strip() for p in re.split(r"\n{2,}", txt) if p.strip()]
    if not paras:
        paras = [txt]

    chunks: List[str] = []
    cur = ""
    for p in paras:
        if not cur:
            cur = p
            continue
        if len(cur) + 2 + len(p) <= max_chunk_chars:
            cur += "\n\n" + p
        else:
            chunks.append(cur)
            cur = p
    if cur:
        chunks.append(cur)

    if len(chunks) <= max_chunks:
        return chunks

    # 앞/중간/뒤를 남겨 긴 문서의 정보 손실 완화
    if max_chunks <= 1:
        return [chunks[0]]
    if max_chunks == 2:
        return [chunks[0], chunks[-1]]

    mid_idx = len(chunks) // 2
    sampled = [chunks[0], chunks[mid_idx], chunks[-1]]
    return sampled[:max_chunks]


# -----------------------------
# 슬라이드 정리
# -----------------------------
# Gemini 생성 슬라이드의 문체/불릿/시각요소 값 정리 함수
def _repair_slides(
    slides: List[Dict[str, Any]],
    *,
    client: genai.Client | None = None,
    model: str = "",
) -> List[Dict[str, Any]]:
    banned = ["본 슬라이드", "추후 보완", "제공되지 않아", "원문 근거 부족"]
    for s in slides:
        s["image_needed"] = False
        s["image_type"] = "none"
        s["image_brief_ko"] = ""

        if _slide_has_formal_lines(s) and client is not None:
            try:
                s = _rewrite_formal_lines_with_gemini(client, model, s)
            except Exception:
                pass

        km = str(s.get("key_message") or "")
        if any(b in km for b in banned):
            s["key_message"] = ""
        s["slide_title"] = _to_phrase(s.get("slide_title"))
        if _contains_formal_line(s["slide_title"]):
            s["slide_title"] = ""

        bullets = s.get("bullets") or []
        nb = []
        if isinstance(bullets, list):
            for b in bullets:
                bt = str(b or "").strip()
                if not bt:
                    continue
                if any(x in bt for x in banned):
                    continue
                bp = _to_phrase(bt)
                if bp and (not _contains_formal_line(bp)):
                    nb.append(bp)
        s["bullets"] = nb
        s["key_message"] = _format_key_message(s.get("key_message"), s.get("slide_title"), s["bullets"])
        if _contains_formal_line(s["key_message"]):
            s["key_message"] = _format_key_message("", s.get("slide_title"), s["bullets"])

        ev_out: List[Dict[str, str]] = []
        for ev in (s.get("evidence") or []):
            if not isinstance(ev, dict):
                continue
            ev_text = _to_phrase(ev.get("text"))
            if not ev_text:
                continue
            if _contains_formal_line(ev_text):
                continue
            ev_out.append({"type": str(ev.get("type") or "근거").strip(), "text": ev_text})
        s["evidence"] = ev_out

        if len(s["bullets"]) < 3:
            s["_drop_slide"] = True

        # "미기재" 도식 강제는 하지 않음(후처리에서 이미지 제거/도식 그리기)
        for k in ["TABLE_MD", "DIAGRAM_SPEC_KO", "CHART_SPEC_KO"]:
            v = str(s.get(k) or "").strip()
            if "미기재" in v or "원문" in v:
                s[k] = ""

    return [s for s in slides if not s.get("_drop_slide")]


# -----------------------------
# LangGraph 노드
# -----------------------------
# 섹션별 Gemini 호출 및 section_decks 생성 함수
def section_deck_generation_node(state: Dict[str, Any]) -> Dict[str, Any]:
    sections = state.get("sections")
    extracted_text = (state.get("extracted_text") or "").strip()

    if not (isinstance(sections, list) and sections):
        if not extracted_text:
            raise RuntimeError("입력 텍스트가 비어 있습니다. (extracted_text/sections)")
        sections = [{"title": (state.get("default_section") or "연구 개요"), "text": extracted_text}]

    client: genai.Client = get_gemini_client()
    prompt = _build_prompt()

    section_decks: Dict[str, Any] = {}
    deck_title = (state.get("deck_title") or "").strip()
    order_cursor = 1

    for s in sections:
        sec_title = re.sub(r"\s+", " ", (s.get("title") or "")).strip()
        sec_text = (s.get("text") or "").strip()

        # 방어적 섹션명 표준화: 정상 흐름에서는 section_split_node의 표준 목차명이 들어오지만,
        # 중간 산출물/수동 입력/이전 산출물 재사용 시 공백·표현 차이를 한 번 더 보정
        alias = {
            "연구내용": "연구 내용",
            "추진계획": "추진 계획",
            "기대효과": "활용방안 및 기대효과",
            "활용계획": "활용방안 및 기대효과",
            "사업개요": "연구 개요",
            "사업 개요": "연구 개요",
            "연구개요": "연구 개요",
            "사업화 계획": "사업화 전략 및 계획",
            "보안조치 이행계획": "사업화 전략 및 계획",
            "안전조치 이행계획": "사업화 전략 및 계획",
        }
        sec_title = alias.get(sec_title, sec_title).strip()

        if not sec_title:
            continue

        # 기관 소개는 서비스 DB 확장 코드 적용 전에도 1장 고정 유지
        if sec_title == "기관 소개":
            one_slide = {
                "order": order_cursor,
                "section": "기관 소개",
                "slide_title": "기관 소개 및 수행역량",
                "key_message": "기관 정보, 수행역량, 참여 역할",
                "bullets": ["기관 소개 정보 입력 필요", "기관 핵심역량 및 수행실적 입력 필요", "주관/참여기관 역할 입력 필요"],
                "evidence": [],
                "image_needed": False,
                "image_type": "none",
                "image_brief_ko": "",
                "TABLE_MD": (
                    "| 항목 | 내용 |\n"
                    "|---|---|\n"
                    "| 기관 소개 | 기관 정보 입력 필요 |\n"
                    "| 수행역량 | 서비스 DB 확장 코드 적용 시 반영 |\n"
                ),
                "DIAGRAM_SPEC_KO": "",
                "CHART_SPEC_KO": "",
            }
            section_decks[sec_title] = {
                "section": sec_title,
                "deck_title": deck_title or "발표자료",
                "slides": [one_slide],
            }
            order_cursor += 1
            continue

        # Q&A는 여기서 만들지 않음(merge에서 강제 추가)
        if sec_title.upper() in {"Q&A", "QNA", "QA"} or sec_title in {"질의응답", "질문", "응답"}:
            continue

        # 너무 짧아도 그냥 넘기지 말고 그대로 보냄(“미기재” 덧붙이지 않음)
        max_chunk_chars = int(state.get("max_section_chunk_chars") or 6000)
        max_chunks = int(state.get("max_section_chunks_per_section") or 3)
        if sec_title == "연구 내용":
            max_chunk_chars = int(state.get("research_content_chunk_chars") or 2400)
            max_chunks = int(state.get("research_content_max_chunks") or 4)
        sec_chunks = _split_section_text_for_llm(
            sec_text,
            max_chunk_chars=max_chunk_chars,
            max_chunks=max_chunks,
        )
        if not sec_chunks:
            continue

        common_rules = """
        [공통 강제 규칙]
        - 모든 출력은 한국어. 영어 문장 금지(고유명사/약어만 예외).
        - 메타 문장(본 슬라이드/추후 보완/제공되지 않아) 절대 금지. 부족하면 '미기재'로만 표기.
        - 목차/표지/챕터/파트 같은 구분용 단독 슬라이드 생성 금지.
        - 각 슬라이드는 TABLE/DIAGRAM/CHART 중 최소 1개를 우선 작성한다.
        - 이미지 생성(IMAGE_NEEDED)은 항상 false.
        - 문장 종결은 메모형(~확보/~예정/~검토/~적용) 사용. '~입니다/~합니다' 금지.
        - 불릿은 설명문 대신 핵심 키워드 중심의 짧은 메모체로 작성.
        - 불릿/키메시지에서 '~이다/~입니다/~합니다/~하였다' 서술형 문장 금지.
        """.strip()


        # 섹션별 추가 규칙(필요 최소만)
        if sec_title == "기관 소개":
            section_rules = """
[추가 규칙]
- '기관 소개' 섹션은 슬라이드 1장만 생성한다.
- TITLE은 '기관 소개 및 수행역량'
""".strip()
        elif sec_title == "연구 내용":
            section_rules = """
[추가 규칙]
- '연구 내용'은 압축 최소화. 원문에 있는 세부 항목을 가능한 한 빠짐없이 반영.
- 서로 다른 세부 주제(예: 데이터, 모델, 보안, 운영, 협력)는 반드시 슬라이드 분리.
- 슬라이드 수를 충분히 확보(최소 5장 이상 권장).
- 각 슬라이드 BULLETS는 4~6개 작성(짧은 메모체), 중복 문장 금지.
- 가능한 경우 TABLE_MD/DIAGRAM_SPEC_KO를 함께 작성해 정보 밀도 확보.
- 일반적인 제목(예: '연구 내용', '핵심 포인트', '세부 정리') 단독 사용 금지.
- 원문의 용어를 유지해 제목/불릿에 구체 키워드(모델명, 데이터, 산출물, 일정)를 반드시 포함.
""".strip()
        elif sec_title == "연구 개요":
            section_rules = """
[추가 규칙]
- '연구 개요'는 최소 1장 이상 생성.
- 과제 개요, 대상기술, 범위/목적을 우선 반영.
- 개요/범위/목적은 설명문 금지, 명사구 bullet로만 작성.
- KEY_MESSAGE는 반드시 키워드 3개(쉼표 구분)로 작성.
- 단일 그림 1개만 있는 슬라이드 금지.
- 2~3개 박스/카드 기반으로 개요·범위·대상기술을 구조적으로 설명.
- 가능하면 TABLE_MD 또는 카드형 비교 구조를 포함.
""".strip()
        elif sec_title == "연구 필요성":
            section_rules = """
[추가 규칙]
- '연구 필요성'은 최소 3장 이상 생성.
- 국내외현황(1-2), 중요성/선행연구/중복성(1-3~1-5)을 분리 반영.
""".strip()
        elif sec_title == "연구 목표":
            section_rules = """
[추가 규칙]
- '연구 목표'는 최소 2장 이상 생성.
- 최종목표와 정량/정성 성능목표를 분리해 작성.
""".strip()
        elif sec_title == "사업화 전략 및 계획":
            section_rules = """
[추가 규칙]
- '사업화 전략 및 계획'은 최소 2장 이상 생성.
- 시장동향/지식재산권/표준화 전략/사업화 계획을 분리해 작성.
""".strip()
        else:
            section_rules = ""

        prompt_for_section = f"{prompt}\n\n{common_rules}\n\n{section_rules}".strip()
        print("[DEBUG][gemini] section:", repr(sec_title), "chunks:", len(sec_chunks), "src_len:", len(sec_text))

        section_slides: List[Dict[str, Any]] = []
        for idx, chunk_text in enumerate(sec_chunks, 1):
            chunk_header = f"[섹션: {sec_title}] [분할 {idx}/{len(sec_chunks)}]\n"
            input_text = chunk_header + chunk_text
            resp = generate_content_with_retry(
                client,
                model=state.get("gemini_model") or "gemini-2.5-flash",
                contents=[prompt_for_section, input_text],
                config=types.GenerateContentConfig(
                    max_output_tokens=int(state.get("gemini_max_output_tokens") or 8192),
                    temperature=float(state.get("gemini_temperature") or 0.4),
                ),
                max_retries=int(state.get("gemini_max_retries") or 5),
            )

            raw = (getattr(resp, "text", None) or "").strip()
            print("[DEBUG][gemini] raw_len:", len(raw), "section:", repr(sec_title), "chunk:", idx)
            if not raw:
                continue

            if not deck_title:
                deck_title = _parse_deck_title(raw).strip()

            slides = _parse_slides_from_text(raw, default_section=sec_title, start_order=order_cursor + len(section_slides))
            slides = _repair_slides(
                slides,
                client=client,
                model=state.get("gemini_model") or "gemini-2.5-flash",
            )
            if not slides:
                slides = _fallback_slide_from_raw(raw, default_section=sec_title, order=order_cursor + len(section_slides))
                slides = _repair_slides(
                    slides,
                    client=client,
                    model=state.get("gemini_model") or "gemini-2.5-flash",
                )
            if slides:
                section_slides.extend(slides)

        if not section_slides:
            continue

        # 섹션 내 중복 제목 최소 제거
        seen_titles = set()
        deduped: List[Dict[str, Any]] = []
        for sl in section_slides:
            key = re.sub(r"\s+", "", str(sl.get("slide_title") or "").lower())
            if key and key in seen_titles:
                continue
            if key:
                seen_titles.add(key)
            deduped.append(sl)
        if not deduped:
            deduped = section_slides

        for i, sl in enumerate(deduped, start=order_cursor):
            sl["order"] = i
        order_cursor += len(deduped)

        section_decks[sec_title] = {
            "section": sec_title,
            "deck_title": deck_title or "발표자료",
            "slides": deduped,
        }

    if not section_decks:
        raise RuntimeError("Gemini가 섹션별 슬라이드를 생성하지 못했습니다. (section_decks=empty)")

    # 제목 후보 미확정 상태 유지: 최종 제목은 merge_deck_node에서 원문/파일명 기준으로 보정
    state["deck_title"] = deck_title or ""
    state["section_decks"] = section_decks

    print("[DEBUG] deck_title:", state["deck_title"])
    print("[DEBUG] section_decks keys:", list(section_decks.keys()))
    total_slides = sum(len(v.get("slides") or []) for v in section_decks.values())
    print("[DEBUG] total slides:", total_slides)

    return state
