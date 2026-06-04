# 국가 R&D 공고 준비 AI 지원 플랫폼

국가 R&D 입찰 준비 과정에서 필요한 반복 업무를 AI로 보조하는 플랫폼입니다.

프로젝트는 사용자가 필요한 기능을 선택해 사용할 수 있는 모듈형 구조로 구성했습니다.

- 공고문 분석
- 자격 체크리스트 제공
- 유사 RFP 탐색 및 제공
- 발표자료 생성
- 발표 대본 및 예상 질문 생성

이 저장소는 그중 제가 담당한 **PDF 텍스트 추출**, **공고문 심층 분석**, **발표자료 생성 파이프라인**을 중심으로 정리했습니다.

## 담당 범위

- PDF 문서 텍스트 추출 및 구조화
- 공고문/RFP 기반 심층 분석 로직 구현
- 제안서 PDF 기반 PPT 초안 생성 파이프라인 구현
- LangGraph 기반 생성 흐름 구성
- Gemini API 응답 구조화 및 파싱
- Gamma API 기반 PPTX 생성
- python-pptx 기반 PPTX 후처리

## 핵심 구현

### 1. PDF 텍스트 추출

공고문과 제안서 PDF에서 텍스트를 추출할 때 단순히 전체 텍스트만 가져오지 않고, 문서 구조를 최대한 유지하는 데 집중했습니다.

`pdfplumber`를 활용해 페이지별로 일반 텍스트, 표, 이미지 영역을 구분해 추출했습니다. 표는 텍스트 기반 표 형태로 변환했고, 페이지 안의 요소들은 위치 정보 기준으로 다시 정렬해 원문 흐름을 최대한 유지하도록 했습니다.

이 파싱 로직은 공고문 분석과 발표자료 생성에서 공통으로 사용했습니다.

### 2. 공고문 심층 분석

공고문 분석에서는 PDF에서 추출한 결과를 바로 Gemini에 전달하지 않고, 문서의 목차 번호와 제목을 기준으로 본문을 섹션 단위로 나누었습니다.

각 섹션은 `chunk_id`와 `text`를 가진 청크 형태로 변환했습니다. 이후 공고문 청크와 RFP 청크를 Gemini에 전달해 사업 배경, 핵심 이슈, 평가항목, 평가항목별 대응 전략, 경쟁력 확보 전략을 JSON 형태로 생성하도록 구성했습니다.

이 기능은 단순 요약보다, 입찰 준비에 필요한 평가 포인트와 제안서 작성 방향을 구조화하는 데 초점을 두었습니다.

### 3. 발표자료 생성 파이프라인

발표자료 생성은 사용자가 업로드한 제안서 PDF를 PPT 초안으로 변환하는 파이프라인입니다.

전체 흐름은 다음과 같습니다.

```text
제안서 PDF 입력
→ 텍스트 추출
→ 발표자료 목차 기준 섹션 분리
→ 섹션별 Gemini 슬라이드 초안 생성
→ 슬라이드 병합 및 순서 정리
→ Gamma API 기반 PPTX 생성
→ python-pptx 후처리
```

Gemini에는 섹션별 원문을 전달하고, 슬라이드 제목, 핵심 메시지, 불릿, 표, 도식 설명 등을 정해진 형식으로 생성하도록 했습니다. 이후 코드에서 응답을 파싱해 섹션별 슬라이드 딕셔너리로 저장했습니다.

생성된 섹션별 슬라이드 초안은 `merge_deck_node.py`에서 하나의 `deck_json`으로 병합했습니다. 병합 단계에서는 표지, 목차, 본문, Q&A 흐름에 맞게 순서를 정리하고, 중복 제목 슬라이드 제거와 최소 장수 보완을 수행했습니다.

마지막으로 Gamma API를 통해 PPTX 파일을 생성하고, python-pptx로 다시 열어 빈 placeholder나 불필요한 이미지 영역 등 형식 문제를 정리했습니다.

## LangGraph 파이프라인

| 노드 | 파일 | 역할 |
| --- | --- | --- |
| `extract_text` | `extract_text_node.py` | PDF/JSON 입력에서 텍스트 추출 |
| `split_sections` | `section_split_node.py` | 발표자료 목차 기준 섹션 분리 |
| `make_sections` | `section_deck_generation_node.py` | 섹션별 Gemini 슬라이드 초안 생성 |
| `merge_deck` | `merge_deck_node.py` | 섹션별 결과를 하나의 `deck_json`으로 병합 |
| `make_pptx` | `gamma_generation_node.py` | Gamma API 호출 및 PPTX 생성 |
| `postprocess` | `postprocess_diagrams.py` | 생성된 PPTX 형식 보정 |

## 기술 스택

- Python
- pdfplumber
- LangGraph
- Gemini API
- Gamma API
- python-pptx
- JSON 기반 응답 파싱

## 문제 해결 포인트

### LLM 결과 불안정성 대응

LLM 결과가 섹션별로 누락되거나 슬라이드 순서가 흔들리는 문제가 있었습니다. 이를 해결하기 위해 문서 전체를 한 번에 처리하지 않고, 발표자료 목차 기준으로 섹션을 나누어 Gemini를 호출했습니다.

또한 Gemini 응답을 자유 형식으로 받지 않고, 슬라이드 제목, 핵심 메시지, 불릿, 표, 도식 설명 같은 항목으로 구조화했습니다. 코드에서는 이 응답을 파싱해 후속 병합 단계에서 사용할 수 있는 데이터로 변환했습니다.

### 병합과 후처리 분리

병합 단계에서는 섹션별 슬라이드 초안을 하나의 `deck_json`으로 합치고, 표지·목차·본문·Q&A 순서에 맞게 발표 흐름을 정리했습니다. 또한 제목만 있거나 내용이 부족한 슬라이드는 제외하고, 섹션별 최소 장수보다 부족한 경우 원문 chunk를 활용해 보완 슬라이드를 추가했습니다.

후처리 단계에서는 Gamma가 생성한 실제 PPTX 파일을 다시 열어 빈 placeholder나 불필요한 이미지 영역처럼 결과물에 남는 형식 문제를 정리했습니다.

## 배운 점

- LLM API를 단순히 연결하는 것만으로는 원하는 결과물이 안정적으로 나오지 않을 수 있다는 점을 경험했습니다.
- 문서 입력을 목적에 맞게 구조화하고, LLM 응답을 코드가 다룰 수 있는 데이터 형태로 정리하는 과정이 중요하다는 점을 배웠습니다.
- 발표자료처럼 목차와 순서가 중요한 산출물은 섹션 분리, 구조화된 응답, 병합, 후처리 단계를 나누어야 결과 흐름을 안정적으로 맞출 수 있었습니다.

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
