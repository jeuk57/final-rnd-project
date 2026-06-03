"""
PDF 파싱 JSON 파일을 목차 기반으로 섹션별로 분리하는 도구
"""

import json
import re
from pathlib import Path
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass, asdict


@dataclass
class Section:
    """문서의 한 섹션을 나타내는 클래스"""
    section_number: str
    title: str
    content: List[str]
    start_page: int
    end_page: int
    
    def to_dict(self):
        return asdict(self)
    
    def to_text(self) -> str:
        """섹션을 텍스트로 변환"""
        text = f"{'='*80}\n"
        text += f"섹션 {self.section_number}: {self.title}\n"
        text += f"페이지: {self.start_page + 1} ~ {self.end_page + 1}\n"
        text += f"{'='*80}\n\n"
        text += "\n".join(self.content)
        return text


class SectionSplitter:
    """PDF 파싱 결과를 목차 기반으로 섹션별로 분리하는 클래스"""
    
    def __init__(self, json_path: str):
        self.json_path = Path(json_path)
        self.pages = self._load_json()
        
        self.toc_patterns = [
            r'^목\s*차$',
            r'^<\s*목\s*차\s*>$',
            r'^contents?$',
            r'^table\s+of\s+contents?$',
            r'^\[목\s*차\]$',
            r'^【목\s*차】$',
        ]
        
        self.section_extraction_patterns = [
            r'^([\d]+(?:[.\-][\d]+)*)[.\s]',
            r'^([가-힣])[.\)]',
            r'^([IVXivx]+)[.\)]',
            r'^([A-Za-z])[.\)]',
        ]
    
    def _load_json(self) -> List[Dict]:
        """JSON 파일 로드"""
        with open(self.json_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def _is_toc_header(self, text: str) -> bool:
        """목차 헤더인지 확인"""
        text_clean = text.strip()
        for pattern in self.toc_patterns:
            if re.match(pattern, text_clean, re.IGNORECASE):
                return True
        return False
    
    def _extract_section_number(self, text: str) -> Optional[str]:
        """
        텍스트에서 섹션 번호만 추출 (제목은 무시)
        """
        text_clean = text.strip()
        # 점선 제거
        text_clean = re.sub(r'[·\s\.]+$', '', text_clean).strip()
        
        for pattern in self.section_extraction_patterns:
            match = re.match(pattern, text_clean)
            if match:
                return match.group(1)
        
        return None
    
    def _extract_number_and_title(self, text: str) -> Tuple[Optional[str], Optional[str]]:
        """
        텍스트에서 섹션 번호와 제목 추출 (목차용)
        """
        text_clean = text.strip()
        text_clean = re.sub(r'[·\s\.]+$', '', text_clean).strip()
        
        for pattern in self.section_extraction_patterns:
            match = re.match(pattern, text_clean)
            if match:
                number = match.group(1)
                title = text_clean[match.end():].strip()
                title = re.sub(r'[·\.]+$', '', title).strip()
                return number, title if title else None
        
        return None, None
    
    def _normalize_text_for_comparison(self, text: str) -> str:
        """텍스트를 정규화하여 비교 (공백, 특수문자 제거)"""
        return re.sub(r'[^\w가-힣]', '', text.lower())
    
    def extract_toc(self) -> Dict[str, str]:
        """
        목차에서 섹션 번호와 제목 추출
        """
        toc = {}
        in_toc = False
        toc_page_idx = None
        
        for page in self.pages:
            page_idx = page['page_index']
            
            # 목차 페이지를 벗어나면 종료
            if in_toc and toc_page_idx is not None and page_idx > toc_page_idx:
                break
            
            for text in page['texts']:
                text_stripped = text.strip()
                
                if self._is_toc_header(text_stripped):
                    in_toc = True
                    toc_page_idx = page_idx
                    continue
                
                if in_toc:
                    if re.match(r'^-?\d+\s*-?$', text_stripped):
                        in_toc = False
                        continue
                    
                    number, title = self._extract_number_and_title(text_stripped)
                    
                    if number and title:
                        toc[number] = title
        
        return toc
    
    def _is_section_match(self, text: str, section_num: str, toc_title: str) -> bool:
        """
        텍스트가 목차의 특정 섹션과 일치하는지 확인
        번호만 비교 (제목은 무시)
        """
        extracted_num = self._extract_section_number(text.strip())
        return extracted_num == section_num
    
    def split_into_sections(self) -> List[Section]:
        """
        문서를 섹션별로 분리
        """
        toc = self.extract_toc()
        
        if not toc:
            print("목차를 찾을 수 없습니다. 전체 문서를 하나의 섹션으로 처리합니다.")
            return [self._create_full_document_section()]
        
        print(f"발견된 목차 항목: {len(toc)}개")
        for num, title in toc.items():
            print(f"   {num}. {title}")
        
        sections = []
        current_section = None
        in_toc_area = False
        toc_page_idx = None
        
        for page in self.pages:
            page_idx = page['page_index']
            
            # 목차 페이지를 벗어나면 목차 영역 종료
            if in_toc_area and toc_page_idx is not None and page_idx > toc_page_idx:
                in_toc_area = False
            
            for text in page['texts']:
                text_stripped = text.strip()
                
                if self._is_toc_header(text_stripped):
                    in_toc_area = True
                    toc_page_idx = page_idx
                    continue
                
                if in_toc_area:
                    if re.match(r'^-?\d+\s*-?$', text_stripped):
                        in_toc_area = False
                    continue
                
                section_started = False
                for section_num, toc_title in toc.items():
                    if self._is_section_match(text_stripped, section_num, toc_title):
                        if current_section:
                            current_section.end_page = page_idx - 1 if page_idx > current_section.start_page else page_idx
                            sections.append(current_section)
                        
                        current_section = Section(
                            section_number=section_num,
                            title=toc_title,
                            content=[],
                            start_page=page_idx,
                            end_page=page_idx
                        )
                        section_started = True
                        break
                
                if section_started:
                    continue
                
                if current_section:
                    if not re.match(r'^-?\d+\s*-?$', text_stripped):
                        current_section.content.append(text)
                        current_section.end_page = page_idx
        
        if current_section:
            sections.append(current_section)
        
        return sections
    
    def _create_full_document_section(self) -> Section:
        """목차가 없을 때 전체 문서를 하나의 섹션으로 생성"""
        content = []
        for page in self.pages:
            content.extend(page['texts'])
        
        return Section(
            section_number="0",
            title="전체 문서",
            content=content,
            start_page=0,
            end_page=len(self.pages) - 1
        )
    
    def save_sections(self, output_path: str, format: str = 'json'):
        """섹션을 파일로 저장"""
        sections = self.split_into_sections()
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        if format == 'json':
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump([s.to_dict() for s in sections], f, ensure_ascii=False, indent=2)
            print(f"\n{len(sections)}개 섹션을 JSON 형식으로 저장: {output_path}")
        
        elif format == 'text':
            for section in sections:
                safe_section_num = re.sub(r'[^\w\-]', '_', section.section_number)
                section_file = output_path.parent / f"{output_path.stem}_section_{safe_section_num}.txt"
                with open(section_file, 'w', encoding='utf-8') as f:
                    f.write(section.to_text())
            print(f"\n{len(sections)}개 섹션을 텍스트 파일로 저장: {output_path.parent}")
        
        elif format == 'combined_text':
            with open(output_path, 'w', encoding='utf-8') as f:
                for section in sections:
                    f.write(section.to_text())
                    f.write("\n\n")
            print(f"\n{len(sections)}개 섹션을 하나의 텍스트 파일로 저장: {output_path}")
        
        return sections


def verify_sections(sections: List[Section], verbose: bool = True):
    """섹션 분리 검증"""
    if verbose:
        print("\n" + "="*80)
        print("섹션 분리 검증 결과")
        print("="*80)
    
    print(f"\n총 섹션 수: {len(sections)}")
    
    if not sections:
        print("오류: 섹션이 하나도 생성되지 않았습니다.")
        return False
    
    if verbose:
        print("\n" + "-"*80)
        print("각 섹션 상세 정보:")
        print("-"*80)
    
    for i, section in enumerate(sections, 1):
        if verbose:
            print(f"\n{i}. 섹션 {section.section_number}: {section.title}")
            print(f"   페이지: {section.start_page + 1} ~ {section.end_page + 1}")
            print(f"   컨텐츠: {len(section.content)}개 항목")
            if len(section.content) > 0:
                preview = '\n'.join(section.content)[:100].replace('\n', ' ')
                print(f"   미리보기: {preview}...")
    
    print("\n" + "="*80)
    print("섹션 분리 완료")
    print("="*80 + "\n")
    
    return True