"""규칙 기반 섹션 분리 및 애매한 블록의 선택적 Gemini 재분류."""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple
import json
import os
import re


SECTION_ORDER = [
    "기관 소개",
    "연구 개요",
    "연구 필요성",
    "연구 목표",
    "연구 내용",
    "추진 계획",
    "활용방안 및 기대효과",
    "사업화 전략 및 계획",
    "Q&A",
]

_HEADING_NUM_RE = re.compile(r"^\s*(\d+)(?:[.\-](\d+))?[.)]?\s*(.+?)\s*$")
_DOT_LEADER_RE = re.compile(r"·\s*·")
_TOC_LINE_RE = re.compile(r"^\s*\d+(?:-\d+)?\.\s*.+·")

_NOISE_HINTS = [
    "연구개발계획서(본문1)",
    "본 서식은 연구개발계획서 본문1",
    "범부처 통합연구지원시스템",
    "제출 시 불필요하며",
    "목 차",
    "< 본문 1 >",
]

_KEYWORDS: Dict[str, List[str]] = {
    "연구 개요": ["연구개발의개요", "개요", "과제개요", "대상기술", "연구범위"],
    "연구 필요성": ["연구개발의필요성", "국내외현황", "중요성", "선행연구", "중복성", "차별성", "필요성", "배경"],
    "연구 목표": ["연구목표", "최종목표", "성과지표", "정량목표", "목표"],
    "연구 내용": ["연구개발과제의내용", "연구내용", "핵심기술", "데이터", "모델", "아키텍처", "구성", "수행일정", "주요결과물"],
    "추진 계획": ["추진전략", "추진방법", "추진체계", "수행체계", "국제공동", "마일스톤", "로드맵", "일정", "추진계획"],
    "활용방안 및 기대효과": ["활용방안", "활용계획", "기대효과", "파급효과", "정책효과", "경제효과", "사회적효과", "성과활용"],
    "사업화 전략 및 계획": ["사업화전략", "사업화계획", "시장동향", "지식재산권", "표준화", "인증기준", "사업화", "안전조치", "보안조치", "이행계획", "보안", "안전"],
}


def _normalize(s: Any) -> str:
    t = str(s or "").replace("\u00a0", " ")
    return re.sub(r"\s+", " ", t).strip()


def _norm_key(s: Any) -> str:
    return re.sub(r"[^0-9a-z가-힣]", "", _normalize(s).lower())


def _parse_heading(line: str) -> Optional[Tuple[int, int, str]]:
    m = _HEADING_NUM_RE.match(_normalize(line))
    if not m:
        return None
    main = int(m.group(1))
    sub = int(m.group(2)) if m.group(2) else 0
    title = _normalize(m.group(3))
    if not (2 <= len(title) <= 120):
        return None
    return main, sub, title


def _section_from_heading(main: int, sub: int, title: str) -> Optional[str]:
    tk = _norm_key(title)

    if main == 1:
        if sub == 1:
            return "연구 개요"
        return "연구 필요성"

    if main == 2:
        if sub in {1, 2}:
            return "연구 목표"
        if sub == 3:
            return "연구 내용"
        if sub == 4:
            return "추진 계획"
        return "연구 목표" if ("목표" in tk) else "연구 내용"

    if main == 3:
        return "추진 계획"

    if main == 4:
        return "활용방안 및 기대효과"

    if main == 5:
        return "사업화 전략 및 계획"

    if main == 6:
        return "사업화 전략 및 계획"

    return None


def _heading_allowed_sections(main: int, sub: int, heading_sec: str) -> List[str]:
    if main == 1:
        if sub == 1:
            return ["연구 개요"]
        return ["연구 필요성"]
    if main == 2:
        if sub in {1, 2}:
            return ["연구 목표"]
        if sub == 3:
            return ["연구 내용"]
        if sub == 4:
            return ["추진 계획"]
        return ["연구 목표", "연구 내용"]
    if main == 3:
        return ["추진 계획"]
    if main == 4:
        return ["활용방안 및 기대효과"]
    if main == 5:
        return ["사업화 전략 및 계획"]
    if main == 6:
        return ["사업화 전략 및 계획"]
    return [heading_sec] if heading_sec else []


def _find_section_headers(lines: List[str]) -> List[Dict[str, Any]]:
    headers: List[Dict[str, Any]] = []
    for i, raw in enumerate(lines):
        parsed = _parse_heading(raw)
        if not parsed:
            continue
        main, sub, title = parsed
        sec = _section_from_heading(main, sub, title)
        if not sec:
            continue
        headers.append({"line_idx": i, "section": sec, "main": main, "sub": sub, "title": title})
    return headers


def _is_noise_line(line: str) -> bool:
    t = _normalize(line)
    if not t:
        return True
    if any(h in t for h in _NOISE_HINTS):
        return True
    if _DOT_LEADER_RE.search(t):
        return True
    if _TOC_LINE_RE.match(t):
        return True
    if t in {"< 본문 1 >", "<본문1>", "목차", "목 차"}:
        return True
    return False


def _clean_chunk(text: str) -> str:
    lines = text.splitlines()
    out: List[str] = []
    for i, raw in enumerate(lines):
        if i > 0 and _is_noise_line(raw):
            continue
        out.append(raw)
    return "\n".join([x for x in out if _normalize(x)]).strip()


def _score_sections(text: str) -> Dict[str, float]:
    tk = _norm_key(text)
    scores: Dict[str, float] = {sec: 0.0 for sec in _KEYWORDS}
    for sec, kws in _KEYWORDS.items():
        for kw in kws:
            k = _norm_key(kw)
            if not k:
                continue
            cnt = tk.count(k)
            if cnt:
                scores[sec] += cnt
                if k in tk[:220]:
                    scores[sec] += 0.7
    return scores


def _best_two(scores: Dict[str, float]) -> Tuple[str, float, str, float]:
    items = sorted(scores.items(), key=lambda x: x[1], reverse=True)
    if not items:
        return "", 0.0, "", 0.0
    if len(items) == 1:
        return items[0][0], items[0][1], "", 0.0
    return items[0][0], items[0][1], items[1][0], items[1][1]


def _is_ambiguous(text: str, heading_section: str) -> Tuple[bool, str, str]:
    scores = _score_sections(text)
    best, best_s, second, second_s = _best_two(scores)
    gap = best_s - second_s

    if best_s < 1.0:
        return True, heading_section, "low_signal"
    if gap < 1.8:
        return True, heading_section, "small_gap"

    hk = _norm_key(heading_section)
    conf = 0.0
    for sec, val in scores.items():
        if _norm_key(sec) == hk:
            conf = val
            break

    if best and best != heading_section and (best_s - conf) >= 2.5 and conf <= 1.0:
        return False, best, "reassign_by_score"

    return False, heading_section, "keep_heading"


def _extract_json_block(text: str) -> str:
    t = (text or "").strip()
    if not t:
        return ""
    i = t.find("{")
    j = t.rfind("}")
    if i >= 0 and j > i:
        return t[i : j + 1]
    return t


def _gemini_reclassify_ambiguous(
    pending: List[Dict[str, Any]],
    state: Dict[str, Any],
) -> Dict[int, str]:
    if not pending:
        return {}

    enabled = state.get("enable_gemini_section_split")
    if enabled is False:
        return {}

    api_key = os.environ.get("GOOGLE_API_KEY", "").strip()
    if not api_key:
        return {}

    try:
        from google import genai
    except Exception:
        print("[WARN][section_split] google.genai not available; skip Gemini reclassify")
        return {}

    model = str(state.get("gemini_model") or "gemini-2.5-flash").strip()
    client = genai.Client(api_key=api_key)

    payload = [
        {
            "id": int(x.get("id")),
            "heading_section": str(x.get("heading_section") or ""),
            "allowed_sections": list(x.get("allowed_sections") or []),
            "text": str(x.get("text") or "")[:2500],
        }
        for x in pending
    ]

    prompt = (
        "너는 국가 R&D 제안서 문서 섹션 분류기다.\n"
        "입력 items의 각 text를 읽고, 해당 item.allowed_sections 중에서만 하나를 고른다.\n"
        "반드시 JSON만 출력한다. 설명 문장 금지.\n"
        "출력 스키마:\n"
        '{"items":[{"id":0,"section":"연구 개요"}]}\n\n'
        f"입력:\n{json.dumps({'items': payload}, ensure_ascii=False)}"
    )

    try:
        resp = client.models.generate_content(model=model, contents=prompt)
        raw = getattr(resp, "text", "") or ""
        data = json.loads(_extract_json_block(raw))
        out: Dict[int, str] = {}

        allow_map = {int(x["id"]): set(x.get("allowed_sections") or []) for x in payload}
        for row in (data.get("items") or []):
            try:
                idx = int(row.get("id"))
            except Exception:
                continue
            sec = _normalize(row.get("section") or "")
            if sec and sec in allow_map.get(idx, set()):
                out[idx] = sec
        if out:
            print(f"[INFO][section_split] Gemini reclassified: {len(out)}/{len(pending)} using {model}")
        return out
    except Exception as e:
        print(f"[WARN][section_split] Gemini reclassify skipped: {e}")
        return {}


def section_split_node(state: Dict[str, Any]) -> Dict[str, Any]:
    extracted_text = state.get("extracted_text") or ""
    lines = (extracted_text or "").splitlines()
    headers = _find_section_headers(lines)

    section_chunks: Dict[str, str] = {sec: "" for sec in SECTION_ORDER}
    debug_rows: List[Dict[str, Any]] = []

    if not headers:
        section_chunks["기관 소개"] = ""
        section_chunks["연구 개요"] = extracted_text
        state["section_chunks"] = section_chunks
        state["sections"] = [{"title": sec, "text": section_chunks[sec]} for sec in SECTION_ORDER]
        state["section_split_debug"] = [{"mode": "no_headers_fallback"}]
        return state

    # 헤더 전 노이즈(prelude)는 기본 제외. 유의미 텍스트만 연구 개요에 편입.
    first_idx = int(headers[0]["line_idx"])
    if first_idx > 0:
        pre = _clean_chunk("\n".join(lines[:first_idx]))
        if len(pre) >= 160:
            section_chunks["연구 개요"] = pre

    section_chunks["기관 소개"] = ""  # 기관 소개는 서비스 DB 확장 코드 적용 시 사용

    pending: List[Dict[str, Any]] = []

    for j, h in enumerate(headers):
        start_idx = int(h["line_idx"])
        heading_sec = str(h["section"])
        main = int(h["main"])
        sub = int(h["sub"])
        end_idx = int(headers[j + 1]["line_idx"]) if (j + 1) < len(headers) else len(lines)

        raw_chunk = "\n".join(lines[start_idx:end_idx]).strip()
        chunk = _clean_chunk(raw_chunk)
        if not chunk:
            continue

        amb, selected_sec, reason = _is_ambiguous(chunk, heading_sec)
        debug_rows.append(
            {
                "line_start": start_idx,
                "main": main,
                "sub": sub,
                "heading_section": heading_sec,
                "selected_section": selected_sec,
                "ambiguous": amb,
                "reason": reason,
                "length": len(chunk),
            }
        )

        if amb:
            allowed_sections = _heading_allowed_sections(main, sub, heading_sec)
            if len(chunk) < 180:
                # 짧은 애매 블록은 헤더 기준으로 유지(LLM 비용/노이즈 방지).
                target = heading_sec
                if section_chunks.get(target):
                    section_chunks[target] += "\n\n" + chunk
                else:
                    section_chunks[target] = chunk
                debug_rows.append(
                    {
                        "line_start": start_idx,
                        "main": main,
                        "sub": sub,
                        "heading_section": heading_sec,
                        "selected_section": target,
                        "ambiguous": False,
                        "reason": "short_ambiguous_keep_heading",
                        "length": len(chunk),
                    }
                )
                continue

            pending.append(
                {
                    "id": len(pending),
                    "heading_section": heading_sec,
                    "allowed_sections": allowed_sections,
                    "text": chunk,
                    "main": main,
                    "sub": sub,
                }
            )
            continue

        if section_chunks.get(selected_sec):
            section_chunks[selected_sec] += "\n\n" + chunk
        else:
            section_chunks[selected_sec] = chunk

    gemini_map = _gemini_reclassify_ambiguous(pending, state)

    for p in pending:
        heading_sec = str(p.get("heading_section") or "")
        chunk = str(p.get("text") or "")
        pid = int(p.get("id"))
        allowed_sections = list(p.get("allowed_sections") or []) or [heading_sec]

        if pid in gemini_map:
            target = gemini_map[pid]
            reason = "gemini_reclassify"
            amb = False
        else:
            target = heading_sec
            reason = "ambiguous_fallback"
            amb = True

        if target not in allowed_sections:
            target = heading_sec
            reason = "allowed_guard_fallback"
            amb = True

        if section_chunks.get(target):
            section_chunks[target] += "\n\n" + chunk
        else:
            section_chunks[target] = chunk

        debug_rows.append(
            {
                "line_start": -1,
                "main": int(p.get("main") or 0),
                "sub": int(p.get("sub") or 0),
                "heading_section": heading_sec,
                "selected_section": target,
                "ambiguous": amb,
                "reason": reason,
                "length": len(chunk),
            }
        )

    state["section_chunks"] = section_chunks
    state["sections"] = [{"title": sec, "text": section_chunks.get(sec, "")} for sec in SECTION_ORDER]
    state["section_split_debug"] = debug_rows
    return state
