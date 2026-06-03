from __future__ import annotations

import os
import re
import time
import json
from datetime import datetime
from typing import Any, Dict, List, Optional
from pathlib import Path

import requests


GAMMA_API_BASE = "https://public-api.gamma.app/v1.0"


# 개발용 재실행 흐름: Gamma 재호출 전 deck_json 중간 산출물 저장 함수
def _save_checkpoint(state: dict) -> str:
    outdir = Path("output") / "checkpoints"
    outdir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = outdir / f"deck_checkpoint_{ts}.json"
    with open(path, "w", encoding="utf-8") as f:
        json.dump(state.get("deck_json", {}), f, ensure_ascii=False, indent=2)
    print(f"[CHECKPOINT] deck_json 저장: {path}")
    return str(path)


# Gamma 입력용 슬라이드 텍스트 변환 함수
def _slides_to_input_text(deck: Dict[str, Any]) -> str:
    title = (deck.get("deck_title") or "").strip() or "발표자료"
    slides: List[Dict[str, Any]] = deck.get("slides") or []
    n = len(slides)

    header = f"""[DECK]
DECK_TITLE: {title}
TOTAL_SLIDES: {n}

절대 규칙:
- 정확히 {n}장만 생성. 추가/삭제/분할/병합 금지.
- 슬라이드 순서 변경 금지.
- 사진/실사/캐릭터/배경 이미지 생성 금지.
- 빈 이미지 placeholder(회색 박스, 깨진 아이콘) 생성 금지.
- 텍스트는 한국어 중심으로 작성(고유명사/약어만 예외).
[/DECK]
""".strip()

    def _strip_formal_endings(text: str) -> str:
        # 종결어미 제거 비활성화: 문장 절단/어색한 마침표 방지
        return str(text or "").strip()

    def _clean_lines(xs: List[str], limit: int) -> List[str]:
        out: List[str] = []
        for x in xs:
            x = _strip_formal_endings(str(x or "")).strip()
            if not x:
                continue
            if x in {"**POST_DIAGRAM_SYSTEM**", "**POST_DIAGRAM_ORGCHART**", "**후처리_대상**"}:
                continue
            out.append(x)
            if len(out) >= limit:
                break
        return out

    slide_blocks: List[str] = []
    for i, s in enumerate(slides, 1):
        section = _strip_formal_endings((s.get("section") or "").strip())
        slide_title = _strip_formal_endings((s.get("slide_title") or "").strip()) or "슬라이드"
        key_message = _strip_formal_endings((s.get("key_message") or "").strip())

        bullets = _clean_lines(s.get("bullets") or [], limit=7)

        table_md = _strip_formal_endings((s.get("TABLE_MD") or "").strip())
        diagram_spec = _strip_formal_endings((s.get("DIAGRAM_SPEC_KO") or "").strip())
        chart_spec = _strip_formal_endings((s.get("CHART_SPEC_KO") or "").strip())

        evidence = s.get("evidence") or []
        ev_lines: List[str] = []
        if isinstance(evidence, list):
            for ev in evidence[:3]:
                if isinstance(ev, dict):
                    t = (ev.get("type") or "근거").strip()
                    tx = _strip_formal_endings((ev.get("text") or "").strip())
                    if tx:
                        ev_lines.append(f"- ({t}) {tx}")
                else:
                    tx = _strip_formal_endings(str(ev or "").strip())
                    if tx:
                        ev_lines.append(f"- {tx}")

        lines: List[str] = []
        lines.append(f"[SLIDE {i}/{n}]")
        lines.append(f"SECTION: {section}")
        lines.append(f"TITLE: {slide_title}")
        if key_message:
            lines.append(f"KEY_MESSAGE: {key_message}")
        lines.append(f"SLIDE_LAYOUT: {(s.get('slide_layout') or '').strip()}")
        lines.append(f"VISUAL_SLOT: {(s.get('visual_slot') or '').strip()}")
        lines.append(f"CONTENT_DENSITY: {(s.get('content_density') or '').strip()}")
        lines.append(f"IMAGE_NEEDED: {bool(s.get('image_needed'))}")
        lines.append(f"IMAGE_TYPE: {(s.get('image_type') or 'none')}")
        if str(s.get("image_brief_ko") or "").strip():
            lines.append(f"IMAGE_BRIEF_KO: {str(s.get('image_brief_ko') or '').strip()}")

        lines.append("BULLETS:")
        if bullets:
            for b in bullets:
                lines.append(f"- {b}")

        if ev_lines:
            lines.append("EVIDENCE:")
            lines.extend(ev_lines)

        if table_md:
            lines.append("TABLE_MD:")
            lines.append(table_md)
        if diagram_spec:
            lines.append("DIAGRAM_SPEC_KO:")
            lines.append(diagram_spec)
        if chart_spec:
            lines.append("CHART_SPEC_KO:")
            lines.append(chart_spec)

        lines.append("[ENDSLIDE]")
        slide_blocks.append("\n".join(lines))

    body = "\n\n---\n\n".join(slide_blocks)
    return header + "\n\n" + body


# Gamma API 요청 헤더 생성 함수
def _gamma_headers(api_key: str) -> Dict[str, str]:
    return {"X-API-KEY": api_key, "Content-Type": "application/json"}


# Gamma 테마 목록 조회 함수
def _list_themes(
    api_key: str,
    *,
    query: str = "",
    limit: int = 50,
    max_pages: int = 5,
) -> List[Dict[str, Any]]:
    themes: List[Dict[str, Any]] = []
    after = ""
    for _ in range(max_pages):
        params: Dict[str, Any] = {"limit": int(limit)}
        if query:
            params["query"] = query
        if after:
            params["after"] = after
        r = requests.get(
            f"{GAMMA_API_BASE}/themes",
            headers=_gamma_headers(api_key),
            params=params,
            timeout=60,
        )
        if r.status_code != 200:
            raise RuntimeError(f"Gamma themes API error {r.status_code}: {r.text}")
        payload = r.json() or {}
        data = payload.get("data") or []
        if isinstance(data, list):
            themes.extend([x for x in data if isinstance(x, dict)])
        if not payload.get("hasMore"):
            break
        after = str(payload.get("nextCursor") or "").strip()
        if not after:
            break
    return themes


# Gamma 테마명/테마 ID 해석 함수
def _resolve_theme_id(api_key: str, theme_input: Optional[str]) -> Optional[str]:
    raw = str(theme_input or "").strip()
    if not raw:
        return None
    # 테마 ID 직접 입력 처리
    if re.fullmatch(r"[A-Za-z0-9_-]{8,}", raw):
        return raw

    themes = _list_themes(api_key, query=raw, limit=50, max_pages=5)
    if not themes:
        return None

    for t in themes:
        if str(t.get("name") or "").strip().lower() == raw.lower():
            return str(t.get("id") or "").strip() or None
    return str((themes[0] or {}).get("id") or "").strip() or None


# Gamma PPTX 생성 작업 시작 요청 함수
def _start_generation(
    api_key: str,
    *,
    input_text: str,
    theme_id: Optional[str],
    num_cards: int,
) -> Dict[str, Any]:

    payload: Dict[str, Any] = {
        "inputText": input_text,
        "format": "presentation",
        "exportAs": "pptx",
        "textMode": "preserve",
        "numCards": int(num_cards),
        "cardOptions": {"dimensions": "16x9"},
        "cardSplit": "inputTextBreaks",

        # 이미지 자동 생성 차단: 빈 이미지 슬롯/회색 박스 생성 방지
        "imageOptions": {"source": "noImages"},

        "textOptions": {
            "language": "ko",
            "tone": "professional, clear",
            "amount": "medium",
        },

        "additionalInstructions": (
            f"정확히 {int(num_cards)}장만 생성. 추가/삭제/분할/병합 금지.\n"
            f"슬라이드 순서 변경 금지.\n"
            f"SECTION 블록 순서 절대 유지: 기관 소개 -> 연구 개요 -> 연구 필요성 -> 연구 목표 -> 연구 내용 -> 추진 계획 -> 활용방안 및 기대효과 -> 사업화 전략 및 계획 -> Q&A.\n"
            f"한 섹션이 시작되면 다음 섹션으로 넘어가기 전까지 해당 섹션 슬라이드를 연속 배치.\n"
            f"영어 문장/영어 제목 금지(고유명사/약어만 예외).\n"
            f"사진/실사/캐릭터/배경 이미지 생성 금지.\n"
            f"TABLE_MD / CHART_SPEC_KO / DIAGRAM_SPEC_KO가 있으면 반드시 반영.\n"
            f"텍스트 밀도 과소 금지: 긴 문단은 금지하되, 슬라이드 당 정보 블록 최소 2개 이상 배치.\n"
            f"설명 문장보다 구조화된 정보 전달(표/도식) 우선.\n"
            f"'추가 정보/문의/연락처' 같은 마무리 슬라이드 생성 금지.\n"
            f"마지막은 '감사합니다' 1장만 허용(중복 금지).\n"
            f"디자인 스타일: 깔끔한 카드형 레이아웃, 균형 배치, 둥근 모서리 중심.\n"
            f"연한 회색 배경 + 블루 포인트 톤. 과도한 빈 공간 금지.\n"
            f"슬라이드별 SLIDE_LAYOUT / VISUAL_SLOT / CONTENT_DENSITY 힌트를 우선 적용.\n"
            f"NotebookLM 스타일처럼 제목-요약-구조화 정보 순서를 유지하고 카드 비율을 일정하게 배치.\n"
            f"한 슬라이드당 핵심 메시지 1개, 불릿은 3~5개 권장.\n"
            f"빈 공간이 크면 카드 2열/요약 박스/표/도식으로 반드시 채운다.\n"
            f"IMAGE_NEEDED=true 인 슬라이드는 빈 이미지 슬롯/회색 박스/깨진 이미지 아이콘을 절대 만들지 않는다.\n"
            f"IMAGE_NEEDED=true 인 슬라이드는 layout=text_image로 생성하고, 우측 40% 영역은 도형/다이어그램/표 등 실제 시각요소로 직접 채운다.\n"
            f"텍스트 박스/도형/표는 이미지/시각요소 영역을 침범하지 않도록 배치(겹침 금지).\n"
            f"시각요소를 생성할 수 없으면 해당 슬라이드를 만들지 말고 이전/다음 슬라이드에 내용 통합.\n"
            f"입력 블록의 TITLE/KEY_MESSAGE/BULLETS는 의미 변경 없이 최대한 원문 그대로 사용.\n"
            f"문장 축약, 재서술, 표현 치환 최소화. 특히 TITLE은 원문 유지.\n"
            f"SLIDE 블록 1개를 카드 1장으로 1:1 매핑하고, 블록 병합/분할 금지.\n"
            f"문장 형태로 작성하지 않는다. 모든 항목은 명사구 또는 키워드 형태로 작성.\n"
            f"문장 종결어미 사용 금지 (~다, ~니다, ~합니다, ~됩니다 포함).\n"
            f"발표 슬라이드용 불릿 형태로 작성. 최소 3개 불릿이 없으면 해당 슬라이드 생성 금지.\n"
            f"내용이 부족하면 슬라이드를 만들지 않는다.\n"
            f"목차 슬라이드에서는 제목만 출력하고 설명 문장은 출력하지 않는다.\n"
            f"표 생성 시 헤더 행 강조 색상, 행별 연한 alternating 색상, 글자색 대비 확보.\n"
            f"표가 필요한 경우 단순한 표 형태 사용.\n"
            f"PowerPoint 표 객체 사용.\n"
            f"카드형 레이아웃 사용 금지.\n"
            f"인포그래픽 스타일 사용 금지.\n"
            f"둥근 카드 장식 사용 금지.\n"
            f"'연구 개요' 섹션은 단독 도식 1장으로만 구성 금지.\n"
            f"2~3개 박스와 보조 불릿을 활용한 구조적 설명 레이아웃 사용.\n"
            f"개념, 범위, 맥락을 설명하는 박스/비교표/매트릭스 구조 우선.\n"
            f"'연구 개요'는 간결한 불릿 문체를 유지하되 설명 정보량 확보.\n"
            f"[SLIDE i/N] ... [ENDSLIDE] 블록 형식 유지."
        ),
    }

    # Gamma 테마 적용: 기본적으로 theme_id가 들어온 경우에만 전달
    if theme_id:
        payload["themeId"] = theme_id

    url = f"{GAMMA_API_BASE}/generations"
    r = requests.post(url, headers=_gamma_headers(api_key), json=payload, timeout=60)
    if r.status_code not in (200, 201):
        raise RuntimeError(f"Gamma API error {r.status_code}: {r.text}")
    return r.json()


# Gamma 생성 작업 완료 상태 조회 함수
def _poll_generation(api_key: str, generation_id: str, *, timeout_sec: int) -> Dict[str, Any]:
    t0 = time.time()
    last: Dict[str, Any] = {}
    while time.time() - t0 < timeout_sec:
        r = requests.get(f"{GAMMA_API_BASE}/generations/{generation_id}", headers=_gamma_headers(api_key), timeout=60)
        r.raise_for_status()
        last = r.json()

        status = (last.get("status") or "").lower()
        if status in {"completed", "complete", "succeeded", "success"}:
            return last
        if status in {"failed", "error"}:
            raise RuntimeError(f"Gamma generation failed: {last}")

        time.sleep(3)

    raise TimeoutError(f"Gamma 생성 대기 시간 초과 ({timeout_sec}s). last={last}")


# Gamma 결과 파일 다운로드 함수
def _download_file(url: str, out_path: str) -> None:
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with requests.get(url, stream=True, timeout=300) as r:
        r.raise_for_status()
        with open(out_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=1024 * 1024):
                if chunk:
                    f.write(chunk)


# Windows 파일 잠금 회피용 출력 경로 보정 함수
def _avoid_windows_lock(path: str) -> str:
    base, ext = os.path.splitext(path)
    if not os.path.exists(path):
        return path
    for i in range(1, 200):
        cand = f"{base} ({i}){ext}"
        if not os.path.exists(cand):
            return cand
    return f"{base}_{int(time.time())}{ext}"


# 파일명 안전 문자 정리 함수
def _safe_filename(name: str) -> str:
    name = re.sub(r"[\\/:*?\"<>|]+", " ", str(name or ""))
    name = re.sub(r"\s+", " ", name).strip()
    if not name:
        return "result"
    # 괄호/기호 제거 후 파일명 길이 축약
    name = re.sub(r"[()\\[\\]{}]", "", name)
    max_len = 36
    if len(name) <= max_len:
        return name
    cut = name[:max_len + 1]
    ws = cut.rfind(" ")
    if ws >= 16:
        return cut[:ws].rstrip()
    return name[:max_len].rstrip()


# Gamma API 기반 PPTX 생성 및 다운로드 노드
def gamma_generation_node(state: Dict[str, Any]) -> Dict[str, Any]:
    api_key = os.environ.get("GAMMA_API_KEY")
    if not api_key:
        raise RuntimeError("GAMMA_API_KEY가 없습니다. .env 또는 환경변수에 설정하세요.")

    deck = state.get("deck_json") or {}
    slides = deck.get("slides") or []
    if not slides:
        raise RuntimeError("deck_json.slides가 비어 있습니다. merge_deck_node 결과를 확인하세요.")

    input_text = _slides_to_input_text(deck)

    output_dir = (state.get("output_dir") or "output").strip()

    # 기본 출력 파일명: 호출자가 지정하지 않으면 안정적인 기본값 사용
    if not (state.get("output_filename") or "").strip():
        output_filename = "RanDi_발표자료.pptx"
    else:
        output_filename = (state.get("output_filename") or "").strip()

    out_path = _avoid_windows_lock(os.path.join(output_dir, output_filename))

    timeout_sec = int(state.get("gamma_timeout_sec") or 600)
    theme_input = (state.get("gamma_theme_id") or state.get("gamma_theme") or "").strip() or None
    theme_id = _resolve_theme_id(api_key, theme_input)
    if theme_input and not theme_id:
        print(f"[WARN] Gamma 테마를 찾지 못했습니다: {theme_input} (themeId 없이 진행)")
    elif theme_id:
        print(f"[INFO] Gamma themeId 확인: {theme_id}")

    if state.get("save_checkpoint", False):
        _save_checkpoint(state)


    gen = _start_generation(api_key, input_text=input_text, theme_id=theme_id, num_cards=len(slides))
    generation_id = gen.get("generationId") or gen.get("id")
    if not generation_id:
        raise RuntimeError(f"Gamma 응답에 generationId가 없습니다: {gen}")

    done = _poll_generation(api_key, generation_id, timeout_sec=timeout_sec)

    def _extract_url(d: Dict[str, Any]) -> str:
        return (
            d.get("exportUrl")
            or d.get("pptxUrl")
            or (d.get("exports") or {}).get("pptx")
            or ""
        )

    file_url = _extract_url(done)

    # completed 직후 다운로드 URL이 늦게 붙는 경우 대비: 최대 45초 추가 조회
    if not file_url:
        t1 = time.time()
        while time.time() - t1 < 45:
            time.sleep(2.5)
            r = requests.get(f"{GAMMA_API_BASE}/generations/{generation_id}", headers=_gamma_headers(api_key), timeout=60)
            r.raise_for_status()
            done2 = r.json()
            file_url = _extract_url(done2)
            if file_url:
                done = done2
                break

    if not file_url:
        raise RuntimeError(f"Gamma 완료 응답에 다운로드 URL이 없습니다: {done}")

    _download_file(file_url, out_path)

    state["final_ppt_path"] = out_path
    state["gamma_ppt_path"] = out_path
    return state
