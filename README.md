# 국가 R&D 공고 준비 AI 지원 플랫폼

국가 R&D 공고문과 제안서 문서를 기반으로 공고 분석과 발표자료 초안 생성을 지원하는 AI 파이프라인입니다.  
이 저장소는 프로젝트 중 제가 담당한 **PDF 텍스트 추출**, **공고문 심층 분석**, **PPT 생성 파이프라인** 중심으로 정리했습니다.

## 담당 범위

- PDF 문서 텍스트 추출 및 문서 구조 보존
- 공고문/RFP 청크 기반 Gemini 심층 분석
- 제안서 PDF 기반 발표자료 생성 파이프라인 설계 및 구현
- LangGraph 기반 단계별 생성 흐름 구성
- Gemini 응답 구조화 및 Gamma API 기반 PPTX 생성
- python-pptx 기반 후처리로 배경, 표지, 이미지/도식 요소 보정

## 주요 기능

### 1. 문서 파싱

`utils/document_parsing.py`에서 PDF와 DOCX 문서를 파싱합니다.  
PDF는 `pdfplumber`를 활용해 일반 텍스트, 표, 이미지 영역을 구분하고, 페이지 내 위치 정보를 기준으로 다시 정렬합니다.  
이 결과를 LLM이 처리하기 쉬운 텍스트 구조로 변환합니다.

### 2. 공고문 심층 분석

`features/rfp_analysis_checklist/notice_llm.py`는 구조화된 공고문/RFP 청크를 Gemini에 전달해 평가항목, 대응 전략 등 심층 분석 결과를 생성하는 로직을 포함합니다.

일부 DB 기반 자격요건 판정 함수는 팀 통합 과정에서 함께 포함되어 있지만, 이 저장소에서는 공고문 심층 분석 흐름을 중심으로 확인하면 됩니다.

### 3. PPT 생성 파이프라인

`features/ppt_maker/main_ppt.py`가 PPT 생성 진입점입니다.

전체 흐름은 다음과 같습니다.

```text
문서 입력
→ 텍스트 추출
→ 발표자료용 섹션 분리
→ 섹션별 Gemini 슬라이드 초안 생성
→ 슬라이드 통합 및 순서 정리
→ Gamma API 기반 PPTX 생성
→ python-pptx 후처리
```

## LangGraph 노드 구조

| 노드 | 파일 | 역할 |
| --- | --- | --- |
| `extract_text` | `extract_text_node.py` | PDF/DOCX/JSON 입력에서 텍스트 추출 |
| `split_sections` | `section_split_node.py` | 발표자료 목차 기준으로 섹션 분리 |
| `make_sections` | `section_deck_generation_node.py` | 섹션별 Gemini 호출로 슬라이드 초안 생성 |
| `merge_deck` | `merge_deck_node.py` | 섹션별 결과를 하나의 `deck_json`으로 병합 |
| `make_pptx` | `gamma_generation_node.py` | Gamma API 호출 및 PPTX 다운로드 |
| `postprocess` | `postprocess_diagrams.py` | 배경, 이미지, 도식 요소 후처리 |

## 기술 스택

- Python
- pdfplumber
- LangGraph
- Gemini API
- Gamma API
- python-pptx
- lxml

## 실행 준비

### 1. 패키지 설치

```bash
pip install -r requirements.txt
```

### 2. 환경변수 설정

`.env.example`을 참고해 `.env` 파일을 생성합니다.

```bash
GOOGLE_API_KEY=
GAMMA_API_KEY=
GAMMA_THEME_ID=
```

실제 API 키가 포함된 `.env` 파일은 GitHub에 올리지 않습니다.

## PPT 생성 실행 예시

```bash
python -m features.ppt_maker.main_ppt --source_path "sample.pdf" --output_dir "outputs"
```

실행 결과는 `output_dir`에 PPTX 파일로 저장됩니다.

## 프로젝트에서 배운 점

- LLM API를 단순히 연결하는 것만으로는 원하는 품질의 결과물이 안정적으로 나오지 않았습니다.
- 문서 입력을 목적에 맞게 구조화하고, 생성 결과를 코드가 다룰 수 있는 중간 데이터로 정리하는 흐름이 중요했습니다.
- 발표자료처럼 목차와 순서가 중요한 산출물은 섹션 분리, JSON 구조화, 병합, 후처리 단계를 나누어야 결과 흐름을 안정적으로 맞출 수 있었습니다.
- 프롬프트 수정만으로 해결하기 어려운 슬라이드 순서, 누락, 형식 문제는 생성 결과를 구조화한 뒤 코드에서 다시 정리하는 방식으로 개선했습니다.

## 디렉터리 구조

```text
features/
  ppt_maker/
    main_ppt.py
    nodes_code/
      extract_text_node.py
      section_split_node.py
      section_deck_generation_node.py
      merge_deck_node.py
      gamma_generation_node.py
      postprocess_diagrams.py
      gemini_diagram_images.py
      llm_utils.py
      state.py
  rfp_analysis_checklist/
    notice_llm.py
    main_notice.py
utils/
  document_parsing.py
  section.py
parsing.py
requirements.txt
```

