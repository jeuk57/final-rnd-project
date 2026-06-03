"""PPT 생성 진입점: 추출 -> 섹션 분리 -> Gemini 섹션 생성 -> 병합 -> 렌더링 -> 후처리."""

import argparse
import json
import os
import re
from datetime import datetime
from typing import Any, Dict, List

from dotenv import load_dotenv
from langgraph.graph import END, START, StateGraph

from features.ppt_maker.nodes_code.extract_text_node import extract_text as extract_text_node
from features.ppt_maker.nodes_code.gamma_generation_node import gamma_generation_node
from features.ppt_maker.nodes_code.merge_deck_node import merge_deck_node
from features.ppt_maker.nodes_code.postprocess_diagrams import postprocess_diagrams_node
from features.ppt_maker.nodes_code.section_deck_generation_node import section_deck_generation_node
from features.ppt_maker.nodes_code.section_split_node import section_split_node
from features.ppt_maker.nodes_code.state import GraphState
from features.ppt_maker.nodes_code.template_render_node import template_render_node
from utils.db_lookup import get_notice_info_by_id

load_dotenv(override=True)

# 개발용 재실행 흐름: Gamma API 미사용 template 모드의 PPTX 템플릿 설정
TEMPLATE_PATH = (os.environ.get("TEMPLATE_PPTX_PATH") or "").strip()
TEMPLATE_LAYOUT_WHITELIST = [
    x.strip() for x in (os.environ.get("TEMPLATE_LAYOUT_WHITELIST") or "").split(",") if x.strip()
]

# 일반 후처리 설정: PPTX 생성 후 postprocess 단계에 전달할 배경 프로필 값
BACKGROUND_IMAGE_PATH = ""
BACKGROUND_PROFILE = "basic"
BACKGROUND_BASE_DIR = os.path.join(os.path.dirname(__file__), "background")
REMOVE_BACKGROUND_IMAGE = False

# 개발용 재실행 흐름: checkpoint deck_json 재사용 시 섹션명과 발표 순서 보정 기준
# 일반 생성 흐름에서는 merge_deck_node.py의 deck_json 정리 결과 사용
CANON_SECTIONS = [
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


# 개발용 재실행 흐름: checkpoint/template 경로 보정에 사용하는 짧은 문자열 공백 정리
def _norm_text(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


# 선택적 DB 확장 코드: notice_id 기반 공고/기관 정보 조회 후 state 주입
def _inject_notice_company_profile(state: Dict[str, Any], notice_id: str) -> None:
    """notice_id가 있으면 DB에서 공고/기관 정보를 조회해 state에 주입."""
    nid = str(notice_id or "").strip()
    if not nid:
        return
    try:
        notice = get_notice_info_by_id(nid)
    except Exception as e:
        print(f"[DB] notice lookup error (notice_id={nid}): {e}")
        return

    if not notice:
        print(f"[DB] no notice found for notice_id={nid}")
        return

    author = str(notice.get("author") or "").strip()
    if author:
        state["org_name"] = author
        state["company_profile"] = {"name": author}
    state["notice_id"] = nid
    state["notice_info"] = notice
    print(f"[DB] notice loaded: notice_id={nid}, org_name={author or '(empty)'}")


# 개발용 재실행 흐름: checkpoint deck_json의 섹션명을 표준 목차명으로 보정
def _canonicalize_section(raw_section: str, slide_title: str) -> str:
    s = _norm_text(raw_section)
    t = _norm_text(slide_title)
    key = f"{s} {t}"

    aliases = {
        "기관소개": "기관 소개",
        "사업 개요": "연구 개요",
        "사업개요": "연구 개요",
        "연구개요": "연구 개요",
        "활용 계획": "활용방안 및 기대효과",
        "기대 효과": "활용방안 및 기대효과",
        "추진계획": "추진 계획",
    }
    s = aliases.get(s, s)
    if s in {"표지", "목차", "Q&A"}:
        return s
    if s in CANON_SECTIONS:
        return s

    if any(k in key for k in ["기관 소개", "수행기관", "주관기관", "참여기관"]):
        return "기관 소개"
    if any(k in key for k in ["연구 개요", "과제 개요", "개요"]):
        return "연구 개요"
    if any(k in key for k in ["연구 필요성", "배경", "필요성", "중요성"]):
        return "연구 필요성"
    if any(k in key for k in ["연구 목표", "최종 목표", "목표", "KPI"]):
        return "연구 목표"
    if any(k in key for k in ["연구 내용", "방법", "모델", "데이터", "아키텍처"]):
        return "연구 내용"
    if any(k in key for k in ["추진 계획", "추진체계", "일정", "마일스톤", "역할"]):
        return "추진 계획"
    if any(k in key for k in ["활용방안", "활용 계획", "기대효과", "성과"]):
        return "활용방안 및 기대효과"
    if any(k in key for k in ["사업화", "시장", "전략", "확산"]):
        return "사업화 전략 및 계획"
    if any(k in key for k in ["Q&A", "질의응답", "질문"]):
        return "Q&A"
    return "연구 내용"


# 개발용 재실행 흐름: checkpoint deck_json 로드 후 섹션명, 목차, 슬라이드 순서 보정
def normalize_and_sort_deck(deck: Dict[str, Any]) -> Dict[str, Any]:
    slides: List[Dict[str, Any]] = deck.get("slides") or []
    if not slides:
        return deck

    for s in slides:
        sec = s.get("section", "")
        title = s.get("slide_title", "")
        s["section"] = _canonicalize_section(sec, title)

        image_type = _norm_text(s.get("image_type", ""))
        brief = _norm_text(s.get("image_brief_ko", ""))
        if any(k in image_type for k in ["사진", "일러스트"]) or any(k in brief for k in ["사진", "유사", "일러스트"]):
            s["image_needed"] = False
        elif not any(k in image_type for k in ["도식", "표", "그래프", "diagram", "block", "system"]):
            s["image_needed"] = False

    cover = [x for x in slides if x.get("section") == "표지"]
    agenda = [x for x in slides if x.get("section") == "목차"]
    rest = [x for x in slides if x.get("section") not in ["표지", "목차"]]

    # 개발용 재실행 흐름: checkpoint deck_json의 기존 슬라이드 order 보존 기준
    def old_order(x: Dict[str, Any]) -> int:
        try:
            return int(x.get("order", 10**9))
        except Exception:
            return 10**9

    section_rank = {sec: i for i, sec in enumerate(CANON_SECTIONS)}
    rest.sort(key=lambda x: (section_rank.get(x.get("section"), 999), old_order(x)))

    merged: List[Dict[str, Any]] = []
    if cover:
        merged.append(cover[0])
    if agenda:
        merged.append(agenda[0])
    merged.extend(rest)

    # 개발용 재실행 흐름: checkpoint deck_json의 목차 슬라이드 항목 보정
    agenda_items = [
        "기관 소개",
        "연구 개요",
        "연구 필요성",
        "연구 목표",
        "연구 내용",
        "추진 계획",
        "활용방안 및 기대효과",
        "사업화 전략 및 계획",
    ]
    if merged and merged[1:2] and merged[1].get("section") == "목차":
        merged[1]["bullets"] = [f"{i+1}. {t}" for i, t in enumerate(agenda_items)]

    for i, s in enumerate(merged, 1):
        s["order"] = i

    deck["slides"] = merged
    return deck


# 개발용 재실행 흐름: checkpoint 경로의 기존 deck_json 로드 및 보정
def _load_deck_checkpoint(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        deck = json.load(f)
    if not isinstance(deck, dict) or not (deck.get("slides") or []):
        raise RuntimeError(f"체크포인트 deck_json이 비어있습니다: {path}")
    return normalize_and_sort_deck(deck)


# 개발용 재실행 흐름: prepare_only 실행 시 deck_json 중간 산출물 파일 저장
def _save_deck_checkpoint(deck: Dict[str, Any], output_dir: str) -> str:
    outdir = os.path.join(output_dir or "output", "checkpoints")
    os.makedirs(outdir, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(outdir, f"deck_prepared_{ts}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(deck or {}, f, ensure_ascii=False, indent=2)
    print(f"[CHECKPOINT] prepared deck_json saved: {path}")
    return path


def build_graph(*, skip_to_gamma: bool = False, prepare_only: bool = False, render_mode: str = "gamma"):
    workflow = StateGraph(GraphState)
    workflow.add_node("make_pptx", gamma_generation_node)
    workflow.add_node("make_template_pptx", template_render_node)
    workflow.add_node("postprocess", postprocess_diagrams_node)

    if skip_to_gamma:
        start_node = "make_template_pptx" if render_mode == "template" else "make_pptx"
        workflow.add_edge(START, start_node)
        workflow.add_edge(start_node, "postprocess")
        workflow.add_edge("postprocess", END)
        return workflow.compile()

    workflow.add_node("extract_text", extract_text_node)
    workflow.add_node("split_sections", section_split_node)
    workflow.add_node("make_sections", section_deck_generation_node)
    workflow.add_node("merge_deck", merge_deck_node)

    workflow.add_edge(START, "extract_text")
    workflow.add_edge("extract_text", "split_sections")
    workflow.add_edge("split_sections", "make_sections")
    workflow.add_edge("make_sections", "merge_deck")
    if prepare_only:
        workflow.add_edge("merge_deck", END)
        return workflow.compile()

    render_node = "make_template_pptx" if render_mode == "template" else "make_pptx"
    workflow.add_edge("merge_deck", render_node)
    workflow.add_edge(render_node, "postprocess")
    workflow.add_edge("postprocess", END)
    return workflow.compile()


def run_ppt_generation(
    *,
    source_path: str = "",
    rfp_text: str = "",
    notice_id: str = "",
    output_dir: str = "",
    output_filename: str = "",
    gemini_model: str = "",
    gamma_theme: str = "cx5kqp1h6rwpfkj",
    gamma_timeout_sec: int = 1800,
    font_name: str = "",
    checkpoint_path: str = "",
    prepare_only: bool = False,
    render_mode: str = "gamma",
):
    print("=" * 80)
    print("PPT 자동 생성 시작 (Extract -> Split -> Gemini -> Merge -> Render)")
    print("=" * 80)

    checkpoint_path = (checkpoint_path or os.environ.get("DECK_CHECKPOINT_PATH") or "").strip()
    skip_to_gamma = bool(checkpoint_path and not prepare_only)

    if (not skip_to_gamma) and (not source_path) and (not rfp_text):
        default_dir = os.path.join(os.getcwd(), "data", "ppt_input")
        if os.path.isdir(default_dir):
            pdfs = [f for f in os.listdir(default_dir) if f.lower().endswith(".pdf")]
            if pdfs:
                source_path = os.path.join(default_dir, pdfs[0])
                print(f"[System] 기본 입력 파일 자동 선택: {source_path}")
            else:
                raise RuntimeError(f"입력 PDF가 없습니다: {default_dir}")
        else:
            raise RuntimeError(f"기본 입력 폴더가 없습니다: {default_dir}")

    render_mode = (render_mode or "gamma").strip().lower()
    if render_mode not in {"gamma", "template"}:
        raise RuntimeError(f"unsupported render_mode: {render_mode}")

    effective_gamma_theme = (gamma_theme or "").strip() or "cx5kqp1h6rwpfkj"
    if BACKGROUND_PROFILE == "brown":
        effective_gamma_theme = os.environ.get("BROWN_GAMMA_THEME_ID") or "ijj5bah3e7ekmcw"
    elif BACKGROUND_PROFILE == "basic":
        effective_gamma_theme = os.environ.get("BASIC_GAMMA_THEME_ID") or effective_gamma_theme

    app = build_graph(skip_to_gamma=skip_to_gamma, prepare_only=prepare_only, render_mode=render_mode)

    initial_state: Dict[str, Any] = {
        # 입력 데이터: 문서 경로 또는 직접 입력 텍스트
        "source_path": source_path,
        "rfp_text": rfp_text,
        # 출력 설정: 결과 PPTX 저장 위치와 파일명
        **({"output_dir": output_dir} if output_dir else {}),
        **({"output_filename": output_filename} if output_filename else {}),
        # 렌더링 설정: Gamma 일반 흐름 또는 template 개발용 재실행 흐름 선택
        "render_mode": render_mode,
        # 개발용 재실행 흐름: template 모드에서 사용할 PPTX 템플릿 설정
        "template_pptx_path": _norm_text(TEMPLATE_PATH),
        "template_layout_whitelist": TEMPLATE_LAYOUT_WHITELIST,
        "template_strict_placeholder_only": True,
        "template_table_as_shape": False,
        # Gemini 생성 설정: 섹션별 슬라이드 초안 생성 모델과 재시도 횟수
        "gemini_model": (gemini_model or "gemini-2.5-flash"),
        "gemini_max_retries": int(os.environ.get("GEMINI_MAX_RETRIES") or 2),
        # Gamma 렌더링 설정: PPTX 생성 테마와 생성 대기 시간
        **({"gamma_theme": effective_gamma_theme} if effective_gamma_theme else {}),
        **({"gamma_timeout_sec": gamma_timeout_sec} if gamma_timeout_sec else {}),
        # 후처리 폰트 설정: 값이 있을 때만 postprocess 단계에 전달
        **({"font_name": font_name} if font_name else {}),
        # 개발용 중간 산출물 설정: deck_json 저장 여부 기본값
        "save_checkpoint": False,
        # Gemini 이미지 생성 설정: 후처리 단계의 이미지 생성/삽입 제어
        "enable_gemini_diagram_images": True,
        "gemini_cover_image_only": False,
        "gemini_image_max_count": 2,
        "gemini_image_model": "models/gemini-2.5-flash-image",
        # 섹션 생성 분량 설정: 긴 섹션 텍스트 분할과 최소 슬라이드 수 제어
        "max_section_chunk_chars": 6000,
        "max_section_chunks_per_section": int(os.environ.get("MAX_SECTION_CHUNKS_PER_SECTION") or 2),
        "min_slide_count": int(os.environ.get("PPT_MIN_SLIDE_COUNT") or 0),
        # 후처리 실행 설정: 표지, 목차, 표, 엔딩, 배경 처리 규칙 제어
        "postprocess_rewrite_cover": True,
        "force_rewrite_cover": True,
        "postprocess_rewrite_agenda": False,
        "postprocess_style_tables": False,
        "postprocess_trim_ending": True,
        "postprocess_apply_template": False,
        "postprocess_apply_background_image": bool(BACKGROUND_IMAGE_PATH or BACKGROUND_PROFILE),
        "postprocess_background_image_path": BACKGROUND_IMAGE_PATH,
        "postprocess_background_profile": BACKGROUND_PROFILE,
        "postprocess_background_base_dir": BACKGROUND_BASE_DIR,
        "postprocess_remove_background_image": REMOVE_BACKGROUND_IMAGE,
        # 결과 데이터 초기값: merge_deck와 렌더링 노드에서 채워질 값
        "deck_json": {},
        "final_ppt_path": "",
    }

    effective_notice_id = str(notice_id or os.environ.get("NOTICE_ID") or "").strip()
    if effective_notice_id:
        _inject_notice_company_profile(initial_state, effective_notice_id)

    if skip_to_gamma:
        loaded_deck = _load_deck_checkpoint(checkpoint_path)
        initial_state["deck_json"] = loaded_deck
        if not source_path and not rfp_text:
            initial_state["source_path"] = ""
        print(f"[System] checkpoint loaded, skip to render: {checkpoint_path}")

    try:
        final_state = app.invoke(initial_state)
        print("\n" + "=" * 80)
        print("PPT 생성 완료")
        print("=" * 80)

        if final_state.get("final_ppt_path"):
            print(f"저장 경로: {final_state['final_ppt_path']}")
        else:
            print("최종 PPT 경로가 비어있습니다. 렌더 단계 실패 가능성이 있습니다.")

        deck = final_state.get("deck_json") or {}
        slides = deck.get("slides") or []
        if prepare_only and deck:
            _save_deck_checkpoint(deck, output_dir or "output")
        print(f"\n슬라이드 수: {len(slides)}")
        if slides:
            print("슬라이드 미리보기 (앞 5개):")
            for i, s in enumerate(slides[:5], 1):
                title = s.get("slide_title") or s.get("title") or "(no title)"
                section = s.get("section") or "(no section)"
                img = s.get("image_needed")
                print(f"  [{i}] [{section}] {title} (image_needed={img})")

        return final_state
    except Exception as e:
        print(f"\n오류 발생: {e}")
        import traceback

        traceback.print_exc()
        return None


def main():
    parser = argparse.ArgumentParser(description="PPT 자동 생성")
    parser.add_argument("--source", default="", help="입력 파일 경로(pdf/docx/json)")
    parser.add_argument("--notice_id", default="", help="DB 확장 코드용 공고 ID (optional)")
    parser.add_argument("--outdir", default="", help="출력 폴더 (default: ./output)")
    parser.add_argument("--outname", default="", help="출력 파일명 (default: result_<id>.pptx)")
    parser.add_argument("--gemini_model", default="", help="Gemini 모델명 (default: gemini-2.5-flash)")
    parser.add_argument("--gamma_theme", default="cx5kqp1h6rwpfkj", help="Gamma 테마 이름 또는 ID")
    parser.add_argument("--gamma_timeout", type=int, default=1800, help="Gamma 생성 상태 확인 제한 시간(초)")
    parser.add_argument("--font_name", default="", help="후처리 폰트명")
    parser.add_argument("--prepare_only", action="store_true", help="렌더 없이 deck_json까지만 생성")
    parser.add_argument("--render_mode", default="gamma", choices=["gamma", "template"], help="최종 렌더 방식")
    parser.add_argument("--checkpoint", default="", help="deck_checkpoint_*.json 경로")
    args = parser.parse_args()

    checkpoint_path = (args.checkpoint or os.environ.get("DECK_CHECKPOINT_PATH") or "").strip()
    required_keys = ["GOOGLE_API_KEY"] if args.prepare_only else ["GOOGLE_API_KEY", "GAMMA_API_KEY"]
    if args.render_mode == "template":
        required_keys = ["GOOGLE_API_KEY"] if not checkpoint_path else []
    if checkpoint_path and not args.prepare_only and args.render_mode == "gamma":
        required_keys = ["GAMMA_API_KEY"]

    missing = [k for k in required_keys if not os.environ.get(k)]
    if missing:
        print(f"환경변수 누락: {', '.join(missing)}")
        print(".env 파일에 API 키를 설정해 주세요.")
        return

    result = run_ppt_generation(
        source_path=args.source,
        notice_id=args.notice_id,
        output_dir=args.outdir,
        output_filename=args.outname,
        gemini_model=args.gemini_model,
        gamma_theme=args.gamma_theme,
        gamma_timeout_sec=args.gamma_timeout,
        font_name=args.font_name,
        checkpoint_path=checkpoint_path,
        prepare_only=args.prepare_only,
        render_mode=args.render_mode,
    )

    if result:
        print("\n작업 완료")
    else:
        print("\n작업 실패")


if __name__ == "__main__":
    main()
