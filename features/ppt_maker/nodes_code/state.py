"""PPT 생성 LangGraph state 스키마."""

from __future__ import annotations

from typing import Any, Dict, List, TypedDict


class GraphState(TypedDict, total=False):
    # 입력값
    source_path: str

    # 텍스트 추출 결과
    extracted_text: str

    # 섹션 분리 결과
    # sections: [{"title": "<섹션명>", "text": "<섹션 텍스트>"} ...]
    sections: List[Dict[str, str]]
    section_chunks: Dict[str, str]

    # 섹션별 Gemini 생성 결과
    # section_decks[section] = {"section":..., "deck_title":..., "slides":[...]}
    section_decks: Dict[str, Any]

    # 병합 결과
    deck_json: Dict[str, Any]
    deck_title: str

    # 출력 설정
    output_dir: str
    output_filename: str
    render_mode: str
    # 개발용 재실행 흐름: 기존 PPTX 템플릿/중간 산출물 기반 렌더링 경로
    template_pptx_path: str
    template_ppt_path: str

    # Gemini 생성 설정
    gemini_model: str
    gemini_temperature: float
    gemini_max_output_tokens: int
    gemini_max_retries: int
    gemini_image_model: str

    # Gamma 생성 설정 및 결과
    gamma_theme: str
    gamma_timeout_sec: int
    gamma_generation_id: str
    gamma_result: Dict[str, Any]
    pptx_url: str
    pptx_path: str

    # 최종 결과
    final_ppt_path: str

    # 선택적 후처리 설정
    font_name: str
    force_rewrite_agenda: bool
    # 개발용 재실행 흐름: deck_json 중간 산출물 저장 여부
    save_checkpoint: bool
    enable_gemini_diagram_images: bool
    gemini_image_max_count: int
    gemini_cover_image_only: bool
    min_slide_count: int
    postprocess_rewrite_cover: bool
    postprocess_rewrite_agenda: bool
    postprocess_style_tables: bool
    postprocess_trim_ending: bool
    postprocess_apply_template: bool
    postprocess_apply_background_image: bool
    postprocess_background_image_path: str
    postprocess_background_profile: str
    postprocess_background_base_dir: str
    postprocess_background_random_seed: int
    postprocess_remove_background_image: bool


def create_empty_state() -> GraphState:
    return {
        "source_path": "",
        "extracted_text": "",
        "sections": [],
        "section_chunks": {},
        "section_decks": {},
        "deck_json": {},
        "deck_title": "",
        "final_ppt_path": "",
    }
