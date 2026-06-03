from __future__ import annotations

import json
import os
import re
from typing import Any, Dict, List

SECTION_ORDER = [
    "기관 소개",
    "연구 개요",
    "연구 필요성",
    "연구 목표",
    "연구 내용",
    "추진 계획",
    "활용방안 및 기대효과",
    "사업화 전략 및 계획",
]

DEFAULT_SECTION_MIN_SLIDES = {
    "기관 소개": 1,
    "연구 개요": 1,
    "연구 필요성": 3,
    "연구 목표": 2,
    "연구 내용": 5,
    "추진 계획": 2,
    "활용방안 및 기대효과": 2,
    "사업화 전략 및 계획": 2,
}

DEFAULT_SECTION_MAX_SLIDES = {
    "사업화 전략 및 계획": 6,
}

IMAGE_KEYWORDS = ["구조", "개요", "흐름", "플랫폼", "서비스", "아키텍처", "시스템"]
IMAGE_DENY_KEYWORDS = ["조직도", "복잡한 구성도", "시장 분석", "예산"]
IMAGE_DENY_SECTIONS = {"연구 목표", "시장 분석", "예산"}


# 텍스트 공백 정규화 함수
def _norm(s: Any) -> str:
    return re.sub(r"\s+", " ", str(s or "").replace("\u00a0", " ")).strip()


# 미기재/불필요 기호 제거 함수
def _clean_text(v: Any) -> str:
    t = _norm(v)
    return t.replace("(미기재)", "").replace("미기재", "").strip(" -:|")


# 발표 슬라이드용 메모체 변환 함수
def _to_memo_phrase(v: Any) -> str:
    t = _clean_text(v)
    if not t:
        return ""
    t = re.sub(r"[.!?]+$", "", t).strip()
    for end in ("입니다", "합니다", "됩니다", "니다", "다"):
        if t.endswith(end):
            t = t[: -len(end)].strip()
            break
    return t


# 원문 텍스트 기반 발표자료 제목 후보 추출 함수
def _extract_title_from_extracted_text(extracted_text: str) -> str:
    text = (extracted_text or "").replace("\u00a0", " ")
    patterns = [
        r"과제명\s*[:·]\s*(.+)",
        r"연구개발\s*과제명\s*[:·]\s*(.+)",
        r"과제\s*제목\s*[:·]\s*(.+)",
        r"사업명\s*[:·]\s*(.+)",
    ]
    for pat in patterns:
        m = re.search(pat, text)
        if m:
            return _norm(m.group(1))[:80]
    # 원문 상단의 유의미한 줄 기반 제목 후보 보완
    ban_phrases = [
        "범부처 통합연구지원시스템",
        "첨부하여 제출",
        "작성하여",
        "제출하여",
        "붙임",
        "유의사항",
        "작성요령",
    ]
    for line in text.splitlines()[:120]:
        ln = _norm(line)
        if not ln:
            continue
        if len(ln) < 8:
            continue
        if re.fullmatch(r"[0-9.\-_/ ]+", ln):
            continue
        if any(bp in ln for bp in ban_phrases):
            continue
        if any(k in ln for k in ["목차", "본문", "요약", "기관 소개", "연구 필요성"]):
            continue
        # 과제명/사업명 접두어 제거
        ln = re.sub(r"^(과제명|사업명|연구과제명)\s*[:：]\s*", "", ln)
        if ln and len(ln) >= 8:
            return ln[:80]
    return ""


# 파일명 기반 발표자료 제목 후보 추출 함수
def _fallback_title_from_filename(state: Dict[str, Any]) -> str:
    src = _norm(state.get("source_path") or "")
    if not src:
        return ""
    base = _norm(os.path.splitext(os.path.basename(src))[0])
    for bad in ["제안서", "사용자업로드", "업로드", "최종", "정식"]:
        base = base.replace(bad, "").strip()
    base = re.sub(r"[_\-]+", " ", base)
    base = re.sub(r"[^\w가-힣 ]+", " ", base)
    base = _norm(base)
    if not base or re.fullmatch(r"[_\-\s]+", base):
        return ""
    return base[:80]


# 발표자료 제목 후보 정제 함수
def _refine_deck_title(title: str) -> str:
    t = _clean_text(title)
    if not t:
        return ""
    ban_phrases = [
        "작성하여 범부처 통합연구지원시스템에 첨부하여 제출",
        "범부처 통합연구지원시스템에 첨부",
        "첨부하여 제출",
        "작성하여 제출",
        "제안서",
    ]
    for bp in ban_phrases:
        t = t.replace(bp, "").strip()
    t = re.sub(r"\s{2,}", " ", t).strip(" -_|")
    if len(t) > 56:
        t = t[:56].rstrip()
    return t


# 섹션별 생성 결과 기반 발표자료 제목 후보 추론 함수
def _infer_title_from_section_decks(section_decks: Dict[str, Any]) -> str:
    preferred_sections = ["연구 개요", "연구 목표", "연구 내용", "연구 필요성"]
    ban = [
        "작성하여",
        "첨부하여 제출",
        "통합연구지원시스템",
        "시행규칙",
        "서식",
        "사업 공고",
        "발표자료",
    ]
    for sec in preferred_sections:
        v = section_decks.get(sec) or {}
        slides = v.get("slides") if isinstance(v, dict) else None
        if not isinstance(slides, list):
            continue
        for s in slides:
            if not isinstance(s, dict):
                continue
            t = _clean_text(s.get("slide_title") or "")
            if len(t) < 8:
                continue
            if any(b in t for b in ban):
                continue
            return t[:56]
    return ""


# 목차 마크다운 표 생성 함수
def _agenda_table(items: List[str]) -> str:
    rows = [f"| {i:02d} | {sec} |" for i, sec in enumerate(items[:8], 1)]
    return "| 번호 | 제목 |\n|---|---|\n" + "\n".join(rows) + "\n"


# 표지 슬라이드 생성 함수
def _make_cover(deck_title: str, org_name: str = "") -> Dict[str, Any]:
    title = _clean_text(deck_title) or "(과제명 미기재)"
    if re.fullmatch(r"[_\-\s]+", title):
        title = "연구개발 과제 제안서"
    table = (
        "| 항목 | 내용 |\n"
        "|---|---|\n"
        f"| 발표 제목 | {title} |\n"
        + (f"| 주관기관 | {_clean_text(org_name)} |\n" if _clean_text(org_name) else "")
    )
    return {
        "order": 1,
        "section": "표지",
        "slide_title": title,
        "key_message": "",
        "bullets": [],
        "evidence": [],
        "image_needed": False,
        "image_type": "none",
        "image_brief_ko": "",
        "TABLE_MD": table,
        "DIAGRAM_SPEC_KO": "",
        "CHART_SPEC_KO": "",
    }


# 일반적이거나 부적절한 제목 후보 판별 함수
def _is_generic_title(title: str) -> bool:
    t = _clean_text(title)
    if not t:
        return True
    if t in {"(과제명 미기재)", "연구개발 과제 제안서"}:
        return True
    if ("개요" in t and "목표" in t) and len(t) <= 20:
        return True
    if re.fullmatch(r"[_\-\s]+", t):
        return True
    return False


# 목차 슬라이드 생성 함수
def _make_agenda() -> Dict[str, Any]:
    return {
        "order": 2,
        "section": "목차",
        "slide_title": "목차",
        "key_message": "",
        "bullets": [],
        "evidence": [],
        "image_needed": False,
        "image_type": "none",
        "image_brief_ko": "",
        "TABLE_MD": _agenda_table(SECTION_ORDER),
        "DIAGRAM_SPEC_KO": "",
        "CHART_SPEC_KO": "",
    }


# 마지막 Q&A/감사 슬라이드 생성 함수
def _make_thanks(order: int, org_name: str = "") -> Dict[str, Any]:
    table = (
        "| 항목 | 내용 |\n"
        "|---|---|\n"
        "| Q&A 진행 | 질의응답 |\n"
        + (f"| 주관기관 | {_clean_text(org_name)} |\n" if _clean_text(org_name) else "")
    )
    return {
        "order": order,
        "section": "Q&A",
        "slide_title": "감사합니다",
        "key_message": "질의응답",
        "bullets": [],
        "evidence": [],
        "image_needed": False,
        "image_type": "none",
        "image_brief_ko": "",
        "TABLE_MD": table,
        "DIAGRAM_SPEC_KO": "",
        "CHART_SPEC_KO": "",
    }


# 섹션별 최소 슬라이드 수 설정 해석 함수
def _resolve_section_min_slides(state: Dict[str, Any]) -> Dict[str, int]:
    mins = {**DEFAULT_SECTION_MIN_SLIDES}
    env_raw = str(os.environ.get("PPT_SECTION_MIN_SLIDES") or "").strip()
    if env_raw:
        try:
            obj = json.loads(env_raw)
            if isinstance(obj, dict):
                for k, v in obj.items():
                    k2 = _norm(k)
                    if k2 in mins:
                        mins[k2] = max(1, int(v))
        except Exception:
            pass
    if isinstance(state.get("section_min_slides"), dict):
        for k, v in state["section_min_slides"].items():
            k2 = _norm(k)
            if k2 in mins:
                try:
                    mins[k2] = max(1, int(v))
                except Exception:
                    pass
    return mins


# 섹션별 최대 슬라이드 수 설정 해석 함수
def _resolve_section_max_slides(state: Dict[str, Any]) -> Dict[str, int]:
    maxs = {sec: 0 for sec in SECTION_ORDER}
    maxs.update(DEFAULT_SECTION_MAX_SLIDES)
    env_raw = str(os.environ.get("PPT_SECTION_MAX_SLIDES") or "").strip()
    if env_raw:
        try:
            obj = json.loads(env_raw)
            if isinstance(obj, dict):
                for k, v in obj.items():
                    k2 = _norm(k)
                    if k2 in maxs:
                        maxs[k2] = max(0, int(v))
        except Exception:
            pass
    if isinstance(state.get("section_max_slides"), dict):
        for k, v in state["section_max_slides"].items():
            k2 = _norm(k)
            if k2 in maxs:
                try:
                    maxs[k2] = max(0, int(v))
                except Exception:
                    pass
    return maxs


# 이미지 대상 후보 슬라이드 판별 함수
def _is_image_candidate(s: Dict[str, Any]) -> bool:
    sec = _clean_text(s.get("section"))
    if sec in IMAGE_DENY_SECTIONS:
        return False
    has_structured = bool(_clean_text(s.get("TABLE_MD")) or _clean_text(s.get("DIAGRAM_SPEC_KO")) or _clean_text(s.get("CHART_SPEC_KO")))
    if has_structured:
        return False
    text_blob = " ".join(
        [
            _clean_text(s.get("slide_title")),
            _clean_text(s.get("key_message")),
            " ".join(_clean_text(x) for x in (s.get("bullets") or [])),
        ]
    )
    if any(k in text_blob for k in IMAGE_DENY_KEYWORDS):
        return False
    return any(k in text_blob for k in IMAGE_KEYWORDS)


# Gamma/후처리용 레이아웃 힌트 부여 함수
def _assign_layout_hints(slide: Dict[str, Any]) -> Dict[str, Any]:
    s = dict(slide)
    sec = _clean_text(s.get("section"))
    has_table = bool(_clean_text(s.get("TABLE_MD")))
    has_diag = bool(_clean_text(s.get("DIAGRAM_SPEC_KO")))
    has_chart = bool(_clean_text(s.get("CHART_SPEC_KO")))
    bullet_count = len(s.get("bullets") or [])

    s["image_needed"] = bool(_is_image_candidate(s))
    s["image_type"] = "diagram" if s["image_needed"] else "none"
    if s["image_needed"] and not _clean_text(s.get("image_brief_ko")):
        s["image_brief_ko"] = "구조/개요/흐름 중심의 발표용 벡터 인포그래픽"

    # 명시적 레이아웃 계약
    if sec in {"표지", "목차", "Q&A"}:
        s["layout"] = "text_only"
    elif s["image_needed"]:
        s["layout"] = "text_image"
    else:
        s["layout"] = "text_only"

    # Gamma 프롬프트 호환용 레이아웃 힌트
    if sec == "표지":
        s["slide_layout"] = "cover"
        s["visual_slot"] = "none"
    elif sec == "목차":
        s["slide_layout"] = "agenda"
        s["visual_slot"] = "none"
    elif has_table and not (has_diag or has_chart):
        s["slide_layout"] = "table_focus"
        s["visual_slot"] = "none"
    elif has_diag or has_chart:
        s["slide_layout"] = "diagram_focus"
        s["visual_slot"] = "right_large"
    elif s["layout"] == "text_image":
        s["slide_layout"] = "text_image"
        s["visual_slot"] = "right_large"
    else:
        s["slide_layout"] = "text_only"
        s["visual_slot"] = "none"

    if bullet_count <= 2:
        s["content_density"] = "low"
    elif bullet_count >= 5:
        s["content_density"] = "high"
    else:
        s["content_density"] = "mid"
    return s


# 유효 슬라이드 판별 및 기본 텍스트 정리 함수
def _is_valid_slide(s: Dict[str, Any]) -> bool:
    bullets = [_to_memo_phrase(b) for b in (s.get("bullets") or []) if _to_memo_phrase(b)]
    s["bullets"] = bullets
    s["slide_title"] = _clean_text(s.get("slide_title"))
    s["key_message"] = _to_memo_phrase(s.get("key_message"))
    # 제목만 있는 슬라이드 금지: 최소 불릿 3개 또는 표/도식 스펙 보유
    return len(bullets) >= 3 or bool(
        _clean_text(s.get("TABLE_MD")) or _clean_text(s.get("DIAGRAM_SPEC_KO")) or _clean_text(s.get("CHART_SPEC_KO"))
    )


# 최소 불릿 수 보장 함수
def _ensure_min_bullets(s: Dict[str, Any], min_count: int = 3) -> Dict[str, Any]:
    out = dict(s)
    bullets = [_to_memo_phrase(b) for b in (out.get("bullets") or []) if _to_memo_phrase(b)]
    key = _to_memo_phrase(out.get("key_message"))
    title = _clean_text(out.get("slide_title"))
    if key and key not in bullets:
        bullets.append(key)
    if title and title not in bullets:
        bullets.append(title)
    extras = [
        "핵심 과업 단계별 수행",
        "주요 산출물 및 검증 지표",
        "리스크 대응 및 협업 체계",
    ]
    for x in extras:
        if len(bullets) >= int(min_count):
            break
        if x not in bullets:
            bullets.append(x)
    out["bullets"] = bullets[: max(int(min_count), 5)]
    return out


# 부족 섹션 보완용 최소 슬라이드 생성 함수
def _make_min_section_slide(section: str, order: int, chunk_text: str = "", variant: int = 1) -> Dict[str, Any]:
    base = _clean_text(chunk_text).replace("\n", " ")
    if len(base) > 120:
        base = base[:120].rstrip() + "..."
    title = f"{section} 핵심 정리" if variant == 1 else f"{section} 핵심 정리 {variant}"
    key = base or f"{section} 핵심 키워드"
    slide = {
        "order": order,
        "section": section,
        "slide_title": title,
        "key_message": key,
        "bullets": [title, key, "발표 핵심 포인트"],
        "evidence": [],
        "image_needed": False,
        "image_type": "none",
        "image_brief_ko": "",
        "TABLE_MD": "",
        "DIAGRAM_SPEC_KO": "",
        "CHART_SPEC_KO": "",
    }
    return _assign_layout_hints(slide)


# 기술 난이도·리스크 대응 고정 슬라이드 생성 함수
def _make_risk_response_slide(section: str, order: int) -> Dict[str, Any]:
    table = (
        "| 난이도 요인 | 리스크 | 대응 |\n"
        "|---|---|---|\n"
        "| 연산 인프라(누리온) | 연산시간·자원부하 | 병렬 최적화, 단계별 실행 |\n"
        "| 장기 관측자료(1990~2025) | 자료 품질 편차·결측 | QC 체계, 표준화 정제 |\n"
        "| 관측·위성 통합(남극/북극, AMSR2/SMOS/CryoSat-2) | 이종 자료 통합 복잡도 | 단계적 개발·검증(1~5단계) |\n"
        "| 자료동화(EnKF)·앙상블 | 구현 난이도·검증 부담 | 모듈 분리, 단계별 성능 검증 |\n"
    )
    return _assign_layout_hints(
        {
            "order": order,
            "section": section,
            "slide_title": "기술 난이도·리스크 및 대응",
            "key_message": "연산 인프라, 장기 관측자료, 통합·검증 난이도",
            "bullets": [
                "대규모 슈퍼컴퓨팅 인프라 필요(KISTI 누리온 활용)",
                "장기 관측자료(1990~2025) 수집·품질관리(QC) 필요",
                "현장 관측(남극/북극)·위성자료(AMSR2/SMOS/CryoSat-2) 활용",
                "자료동화(EnKF), 앙상블 구현 난이도 요소",
            ],
            "evidence": [
                {"type": "출처", "text": "제안서_사용자업로드"},
            ],
            "image_needed": False,
            "image_type": "none",
            "image_brief_ko": "",
            "TABLE_MD": table,
            "DIAGRAM_SPEC_KO": "",
            "CHART_SPEC_KO": "",
        }
    )


# 연구 개요 보강용 기관역량 슬라이드 생성 함수
def _make_why_us_slide(section: str, order: int) -> Dict[str, Any]:
    table = (
        "| 성과 기반 | 데이터·인프라 기반 | 추진체계 기반 |\n"
        "|---|---|---|\n"
        "| 해양순환 모델 고도화(예측 정확도 30% 향상) | 극지 관측 기지(장보고/세종/다산) | 주관·공동·위탁·협력기관 역할분담 |\n"
        "| 극지 해빙 관측·모델 기반 연구 | 장기 관측자료, 위성자료 활용 | 제안서 명시 역할 중심 수행 |\n"
        "| 생태계 영향 평가 경험 | 관측·연산 연계 기반 | 단계별 개발·검증 체계 |\n"
    )
    return _assign_layout_hints(
        {
            "order": order,
            "section": section,
            "slide_title": "왜 우리가 해야 하는가(기관역량·수행근거)",
            "key_message": "성과 기반, 데이터·인프라 기반, 추진체계 기반",
            "bullets": [
                "선행연구 실적: 해양순환 모델 고도화(예측 정확도 30% 향상)",
                "관측·인프라: 장보고/세종/다산 기지, 장기 관측자료·위성자료",
                "추진체계: 주관·공동·위탁·협력기관 역할분담",
            ],
            "evidence": [
                {"type": "출처", "text": "제안서_사용자업로드"},
            ],
            "image_needed": False,
            "image_type": "none",
            "image_brief_ko": "",
            "TABLE_MD": table,
            "DIAGRAM_SPEC_KO": "",
            "CHART_SPEC_KO": "",
        }
    )


# 시스템 아키텍처 고정 슬라이드 생성 함수
def _make_system_architecture_slide(section: str, order: int) -> Dict[str, Any]:
    return _assign_layout_hints(
        {
            "order": order,
            "section": section,
            "slide_title": "시스템 아키텍처",
            "key_message": "",
            "bullets": [],
            "evidence": [{"type": "출처", "text": "제안서_사용자업로드"}],
            "image_needed": True,
            "image_type": "diagram",
            "image_prompt_type": "system_architecture",
            "image_title_only": True,
            "image_brief_ko": "통합 시스템 아키텍처 구성도",
            "TABLE_MD": "",
            "DIAGRAM_SPEC_KO": "",
            "CHART_SPEC_KO": "",
        }
    )


# 기관 소개 임시 문구 판별 함수: 실제 DB 조회 없음, 서비스 DB 확장 코드 대기 문구만 감지
def _is_org_placeholder_slide(slide: Dict[str, Any]) -> bool:
    text_blob = " ".join(
        [
            str(slide.get("slide_title") or ""),
            str(slide.get("key_message") or ""),
            " ".join(str(x or "") for x in (slide.get("bullets") or [])),
        ]
    ).lower()
    return ("db" in text_blob) or ("연동" in text_blob) or ("대기" in text_blob) or ("자동 반영" in text_blob)


# 기관 소개 임시 슬라이드의 근거형 내용 대체 함수
def _rewrite_org_slide_with_evidence(slide: Dict[str, Any]) -> Dict[str, Any]:
    s = dict(slide)
    s["slide_title"] = "기관 소개 및 수행역량"
    s["key_message"] = "선행연구 성과, 관측·연산 인프라, 참여기관 역할분담"
    s["bullets"] = [
        "선행연구 성과 요약: 해양순환 모델 고도화, 극지 해빙 관측·모델 기반, 생태계 영향 평가",
        "관측·연산 인프라: 장보고/세종/다산 기지, 장기 관측자료, 위성자료",
        "참여기관 역할분담: 주관·공동·위탁·협력기관 역할",
    ]
    s["evidence"] = [{"type": "출처", "text": "제안서_사용자업로드"}]
    s["TABLE_MD"] = (
        "| 구분 | 핵심 근거 |\n"
        "|---|---|\n"
        "| 선행연구 성과 | 해양순환 모델 고도화(예측 정확도 30% 향상), 극지 해빙·생태계 영향 평가 |\n"
        "| 관측·연산 인프라 | 장보고/세종/다산 기지, 장기 관측자료, 위성자료 |\n"
        "| 참여기관 역할분담 | 주관·공동·위탁·협력기관 역할 |\n"
    )
    s["DIAGRAM_SPEC_KO"] = ""
    s["CHART_SPEC_KO"] = ""
    s["image_needed"] = False
    s["image_type"] = "none"
    s["image_brief_ko"] = ""
    return _assign_layout_hints(s)


# 이미지 삽입 대상 슬라이드 고정 함수
def _force_fixed_image_targets(slides: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    # 고정 정책:
    # - 표지: 후처리 이미지 노드에서 표지 이미지 영역 처리
    # - 연구 개요 마지막/추진 계획 마지막/시스템 아키텍처 슬라이드만 이미지 대상 지정
    if not slides:
        return slides

    for s in slides:
        if not isinstance(s, dict):
            continue
        s["image_needed"] = False
        s["image_type"] = "none"
        if str(s.get("layout") or "").lower() == "text_image":
            s["layout"] = "text_only"
            s["slide_layout"] = "text_only"
            s["visual_slot"] = "none"

    # 고정 이미지 대상:
    # 1) '연구 개요' 마지막 슬라이드
    # 2) '추진 계획' 마지막 슬라이드
    # 3) '연구 내용'의 '시스템 아키텍처' 슬라이드
    overview_sec = SECTION_ORDER[1] if len(SECTION_ORDER) > 1 else ""
    plan_sec = SECTION_ORDER[5] if len(SECTION_ORDER) > 5 else ""
    content_sec = SECTION_ORDER[4] if len(SECTION_ORDER) > 4 else ""
    overview_last_idx = -1
    plan_last_idx = -1
    for i, s in enumerate(slides):
        if not isinstance(s, dict):
            continue
        if str(s.get("section") or "") == overview_sec:
            overview_last_idx = i
        if str(s.get("section") or "") == plan_sec:
            plan_last_idx = i

    for i, s in enumerate(slides):
        if not isinstance(s, dict):
            continue
        sec = str(s.get("section") or "")
        title = str(s.get("slide_title") or "")
        if i == overview_last_idx and overview_last_idx >= 0:
            s["image_needed"] = True
            s["image_type"] = "diagram"
            s["image_prompt_type"] = "overview_last"
            s["layout"] = "text_image"
            s["slide_layout"] = "text_image"
            s["visual_slot"] = "right_large"
            continue
        if i == plan_last_idx and plan_last_idx >= 0:
            s["image_needed"] = True
            s["image_type"] = "diagram"
            s["image_prompt_type"] = "plan_orgchart_fixed"
            s["layout"] = "text_image"
            s["slide_layout"] = "text_image"
            s["visual_slot"] = "right_large"
            continue
        if sec == content_sec and ("시스템 아키텍처" in title):
            s["image_needed"] = True
            s["image_type"] = "diagram"
            s["image_prompt_type"] = "system_architecture"
            s["image_title_only"] = True
            s["layout"] = "text_image"
            s["slide_layout"] = "text_image"
            s["visual_slot"] = "right_large"
            continue

    return slides


def merge_deck_node(state: Dict[str, Any]) -> Dict[str, Any]:
    default_cover_title = "전지구 해양·극지 고정밀 기후예측시스템 개발"
    deck_title = _norm(state.get("deck_title") or "")
    if not deck_title:
        deck_title = default_cover_title
    if (not deck_title) or ("미기재" in deck_title):
        deck_title = _extract_title_from_extracted_text(state.get("extracted_text") or "") or _fallback_title_from_filename(state) or "연구개발 과제 제안서"
    if _is_generic_title(deck_title):
        guessed = _extract_title_from_extracted_text(state.get("extracted_text") or "")
        if guessed and (not _is_generic_title(guessed)):
            deck_title = guessed
        else:
            fn_title = _fallback_title_from_filename(state)
            if fn_title and (not _is_generic_title(fn_title)):
                deck_title = fn_title
            elif _is_generic_title(deck_title):
                deck_title = "연구개발 과제 제안서"
    refined = _refine_deck_title(deck_title)
    if refined:
        deck_title = refined
    bad_title_markers = [
        "첨부하여 제출",
        "통합연구지원시스템",
        "작성하여 제출",
        "시행규칙",
        "서식",
        "사업 공고",
        "발표자료",
    ]
    section_decks = state.get("section_decks") or {}
    if (not deck_title) or any(m in deck_title for m in bad_title_markers):
        deck_title = default_cover_title

    normalized = {_norm(k): v for k, v in section_decks.items() if _norm(k)}
    section_min = _resolve_section_min_slides(state)
    section_max = _resolve_section_max_slides(state)
    org_name = _norm((state.get("company_profile") or {}).get("name") or state.get("org_name") or "")

    slides: List[Dict[str, Any]] = [_make_cover(deck_title, org_name), _make_agenda()]
    order = 3

    for sec in SECTION_ORDER:
        raw_slides: List[Dict[str, Any]] = []
        v = normalized.get(sec)
        if isinstance(v, dict):
            raw_slides = v.get("slides") or []

        valid: List[Dict[str, Any]] = []
        seen_titles = set()
        for s in raw_slides:
            if not isinstance(s, dict):
                continue
            s2 = dict(s)
            s2["section"] = sec
            if not _is_valid_slide(s2):
                continue
            key = re.sub(r"\s+", "", _clean_text(s2.get("slide_title")).lower())
            if key and key in seen_titles:
                continue
            if key:
                seen_titles.add(key)
            valid.append(s2)

        sec_max = int(section_max.get(sec, 0) or 0)
        if sec_max > 0 and len(valid) > sec_max:
            valid = valid[:sec_max]

        if not valid:
            valid = [_make_min_section_slide(sec, order, chunk_text=str((state.get("section_chunks") or {}).get(sec) or ""), variant=1)]

        # 기관 소개: 플레이스홀더 문구 제거 + 근거형 내용으로 대체
        if sec == SECTION_ORDER[0]:
            patched: List[Dict[str, Any]] = []
            replaced = False
            for s in valid:
                if _is_org_placeholder_slide(s):
                    patched.append(_rewrite_org_slide_with_evidence(s))
                    replaced = True
                else:
                    patched.append(s)
            valid = patched
            if not replaced and valid:
                valid[0] = _rewrite_org_slide_with_evidence(valid[0])

        sec_min = max(1, int(section_min.get(sec, 1)))
        while len(valid) < sec_min:
            valid.append(
                _make_min_section_slide(
                    sec,
                    order + len(valid),
                    chunk_text=str((state.get("section_chunks") or {}).get(sec) or ""),
                    variant=len(valid) + 1,
                )
            )

        # 추진 계획 섹션 시작 직전: 난이도·리스크·대응 1장 추가
        if sec == SECTION_ORDER[5]:
            risk_slide = _make_risk_response_slide(sec, order)
            risk_slide["order"] = order
            order += 1
            slides.append(risk_slide)

        for s in valid:
            s["order"] = order
            order += 1
            if sec == SECTION_ORDER[5]:
                s = _ensure_min_bullets(s, min_count=3)
            slides.append(_assign_layout_hints(s))

        # 연구 내용 섹션 마지막: 시스템 아키텍처 전용 슬라이드 추가
        if sec == SECTION_ORDER[4]:
            arch_slide = _make_system_architecture_slide(sec, order)
            arch_slide["order"] = order
            order += 1
            slides.append(arch_slide)

        # 연구 개요 섹션 다음: "왜 우리가 해야 하는가" 1장 추가
        if sec == SECTION_ORDER[1]:
            why_slide = _make_why_us_slide(sec, order)
            why_slide["order"] = order
            order += 1
            slides.append(why_slide)


    slides.append(_make_thanks(order, org_name))
    slides = _force_fixed_image_targets(slides)
    state["deck_title"] = deck_title
    state["deck_json"] = {"deck_title": deck_title, "slides": slides}
    return state
