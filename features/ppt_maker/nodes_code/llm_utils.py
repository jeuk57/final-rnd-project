"""
Gemini LLM 호출 공통 유틸.

- GOOGLE_API_KEY만 사용 (Gemini 호출)
- 429/5xx 계열 오류의 재시도 및 지수 대기 처리
- 429 응답에 "retry in XXs"가 있으면 해당 시간만큼 대기 후 재시도
"""

from __future__ import annotations

import os
import re
import time
from typing import Any, Optional

from google import genai
from google.genai import types


# Gemini API 키 조회 함수
def get_api_key() -> str:
    api_key = os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        raise RuntimeError("GOOGLE_API_KEY 환경변수가 필요합니다.")
    return api_key


# Gemini 클라이언트 생성 함수
def get_gemini_client() -> genai.Client:
    return genai.Client(api_key=get_api_key())


# Gemini rate limit 메시지의 재시도 대기 시간 추출 함수
def _extract_retry_seconds(msg: str) -> Optional[int]:
    """
    에러 메시지에 'Please retry in 46.7s' / 'retry in 46s' 같은 문구가 있으면 초 단위로 추출
    """
    m = re.search(r"retry in\s+(\d+)(?:\.\d+)?s", msg.lower())
    if m:
        return int(m.group(1))
    return None


# 재시도 불가능한 무료 할당량 차단 여부 판별 함수
def _is_permanent_free_tier_block(msg: str) -> bool:
    """
    재시도해도 해소되지 않는 영구 차단성 quota 오류만 True.
    'limit: 5 ... retry in 46s' 같은 일시적 rate limit은 재시도 대상으로 유지.
    """
    low = msg.lower()
    return ("limit: 0" in low) or ("quotavalue': '0" in low) or ("quota value: 0" in low)


# Gemini 호출 재시도 및 대기 처리 함수
def generate_content_with_retry(
    client: genai.Client,
    *,
    model: str,
    contents: Any,
    config: Optional[types.GenerateContentConfig] = None,
    max_retries: int = 5,
    base_sleep_sec: float = 1.5,
) -> Any:
    last_exc: Optional[Exception] = None

    for attempt in range(max_retries):
        try:
            return client.models.generate_content(model=model, contents=contents, config=config)
        except Exception as e:
            last_exc = e
            msg = str(e)

            # 정말로 0 한도면 즉시 중단
            if _is_permanent_free_tier_block(msg):
                raise RuntimeError(
                    "Gemini free-tier quota가 0(또는 결제/권한 문제로 영구 차단)입니다. "
                    "Billing 연결 또는 프로젝트/키를 확인하세요."
                ) from e

            # 응답 메시지에 retry in이 있으면 해당 시간만큼 대기
            retry_sec = _extract_retry_seconds(msg)
            if retry_sec is not None:
                wait = min(retry_sec + 1, 120)
                print(f"[WARN] Gemini rate limit. {wait}초 대기 후 재시도")
                time.sleep(wait)
                continue

            # 그 외 오류는 지수 대기 방식으로 재시도
            sleep_sec = min(base_sleep_sec * (2 ** attempt), 30.0)
            print(f"[WARN] Gemini 오류. {sleep_sec:.1f}초 대기 후 재시도 ({attempt+1}/{max_retries})")
            time.sleep(sleep_sec)

    raise RuntimeError(f"Gemini 재시도 초과: {last_exc}") from last_exc


# Gamma API 키 조회 함수
def get_gamma_api_key() -> str:
    api_key = os.environ.get("GAMMA_API_KEY")
    if not api_key:
        raise RuntimeError("GAMMA_API_KEY 환경변수가 필요합니다.")
    return api_key
