오프라인 환경에서 웹 에디터(프론트엔드)와 파이썬 백엔드 간에 데이터를 안전하게 주고받기 위한 1) JSON 데이터 규격과 2) 표(Table) 구조 분석 및 예외 처리가 강화된 소스코드를 추가하여 마크다운 파일 내용을 업데이트했습니다.
------------------------------

# 📦 오프라인 통합 문서 변환 시스템 가이드 (PPTX / DOCX ⇄ HTML)
인터넷 연결이 불가능한 폐쇄망 환경에서 대형 AI 모델(Docling 등)이나 유료 라이브러리(Aspose) 없이, **순수 파이썬 초경량 패키지만을 활용하여 이미지, 표, 좌표를 보존하는 양방향 변환 시스템** 가이드입니다.
---## 1. ⚙️ 오프라인 환경 구축 (패키지 사전 준비)
인터넷이 되는 외부 환경에서 필요한 필수 패키지(`.whl`)를 미리 다운로드하여 오프라인 서버로 반입해야 합니다.
### 1) 외부망 PC에서 패키지 다운로드```bash
mkdir offline_packages
cd offline_packages

# 필수 패키지 한 번에 다운로드
pip download langgraph python-pptx python-docx mammoth beautifulsoup4
```
### 2) 폐쇄망 서버에서 패키지 설치다운로드한 폴더를 오프라인 서버로 이동한 후 로컬 경로 지정을 통해 설치합니다.```bash
pip install --no-index --find-links=/path/to/offline_packages langgraph python-pptx python-docx mammoth beautifulsoup4
```
---## 2. 📑 프론트엔드-백엔드 간 JSON 데이터 통신 규격
사용자가 웹 화면에서 슬라이드나 문서를 편집하고 저장 버튼을 눌렀을 때, 오프라인 서버(백엔드)로 전송해야 하는 표준 JSON 데이터 포맷입니다. 
### 1) PPTX 편집 데이터 전송 규격 (슬라이드 형태)```json
{
  "file_type": "pptx",
  "direction": "to_office",
  "original_file_name": "quarterly_report.pptx",
  "slides": [
    {
      "slide_number": 1,
      "elements": [
        {
          "type": "text",
          "content": "수정된 메인 제목입니다",
          "style": { "left": 100, "top": 150, "width": 600, "height": 80 }
        },
        {
          "type": "image",
          "content": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAA...",
          "style": { "left": 150, "top": 250, "width": 400, "height": 300 }
        }
      ]
    }
  ]
}
```
### 2) DOCX 편집 데이터 전송 규격 (스크롤 문서 형태)```json
{
  "file_type": "docx",
  "direction": "to_office",
  "original_file_name": "agreement.docx",
  "document_body": [
    { "type": "h1", "content": "제 1조 목적" },
    { "type": "p", "content": "본 계약은 갑과 을의 권리와 의무를 규정함을 목적으로 합니다." },
    {
      "type": "table",
      "rows": [
        ["구분", "내용", "비고"],
        ["갑", "홍길동", "공급자"],
        ["을", "이순신", "수급자"]
      ]
    }
  ]
}
```
---
## 3. 🛠️ 표(Table) 인식 및 예외 처리가 강화된 통합 모듈 (`converter.py`)

기존 코드에서 누락되기 쉬운 **표(Table) 파싱 기능**을 추가하고, 파일이 손상되었거나 좌표 정규식 매칭이 실패했을 때 시스템이 멈추지 않도록 **Try-Catch 예외 처리 메커니즘**을 강화한 코드입니다.

```python
import os
import re
import base64
import logging
import mammoth
from io import BytesIO
from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches
from bs4 import BeautifulSoup

# 오프라인 로그 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class OfflineConverter:
    """인터넷 없이 로컬에서 구동되는 이미지, 표, 좌표 포함 오피스 양방향 변환 모듈"""
    
    # =============================================================
    # 📊 [파워포인트] PPTX ⇄ HTML (절대좌표, 이미지 및 예외 처리)
    # =============================================================
    @staticmethod
    def pptx_to_html(pptx_path: str, output_html_path: str) -> bool:
        try:
            if not os.path.exists(pptx_path):
                raise FileNotFoundError(f"원본 파일을 찾을 수 없습니다: {pptx_path}")
                
            prs = Presentation(pptx_path)
            sw, sh = int(prs.slide_width / 9525), int(prs.slide_height / 9525)
            
            html = f"<html><head><style>.slide {{ width:{sw}px; height:{sh}px; position:relative; background:white; border:1px solid #ccc; margin-bottom:20px; }} .element {{ position:absolute; }}</style></head><body>"
            
            for i, slide in enumerate(prs.slides):
                html += f"<div class='slide' id='slide-{i+1}'>"
                for shape in slide.shapes:
                    try:
                        l, t = int(shape.left / 9525), int(shape.top / 9525)
                        w, h = int(shape.width / 9525), int(shape.height / 9525)
                        style = f"left:{l}px; top:{t}px; width:{w}px; height:{h}px;"
                        
                        # 1. 텍스트 박스 처리
                        if hasattr(shape, "text") and shape.text.strip() and shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                            html += f"<div class='element' style='{style}'>{shape.text}</div>"
                        
                        # 2. 이미지 처리
                        elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                            base64_str = base64.b64encode(shape.image.blob).decode('utf-8')
                            img_src = f"data:{shape.image.content_type};base64,{base64_str}"
                            html += f"<img class='element' style='{style}' src='{img_src}' />"
                            
                    except Exception as e:
                        logging.warning(f"슬라이드 {i+1}의 특정 도형 파싱 실패(스킵됨): {e}")
                        continue
                html += "</div>"
            html += "</body></html>"
            
            with open(output_html_path, "w", encoding="utf-8") as f:
                f.write(html)
            return True
            
        except Exception as e:
            logging.error(f"PPTX to HTML 변환 중 치명적 오류 발생: {e}")
            return False

    @staticmethod
    def html_to_pptx(html_path: str, output_pptx_path: str) -> bool:
        try:
            prs = Presentation()
            blank_layout = prs.slide_layouts[6] # 완전히 비어있는 레이아웃 안전 선택
            
            with open(html_path, "r", encoding="utf-8") as f:
                soup = BeautifulSoup(f.read(), "html.parser")
                
            for s in soup.find_all(class_='slide'):
                slide = prs.slides.add_slide(blank_layout)
                for elem in s.find_all(class_='element'):
                    try:
                        style_str = elem.get('style', '')
                        # 정규식 예외 처리 강화 (좌표 누락 대비 기본값 0 처리)
                        l_match = re.search(r'left:\s*(\d+)px', style_str)
                        t_match = re.search(r'top:\s*(\d+)px', style_str)
                        w_match = re.search(r'width:\s*(\d+)px', style_str)
                        h_match = re.search(r'height:\s*(\d+)px', style_str)
                        
                        l_px = int(l_match.group(1)) if l_match else 0
                        t_px = int(t_match.group(1)) if t_match else 0
                        w_px = int(w_match.group(1)) if w_match else 100
                        h_px = int(h_match.group(1)) if h_match else 50
                        
                        if elem.name == 'div':
                            txBox = slide.shapes.add_textbox(Inches(l_px/96), Inches(t_px/96), Inches(w_px/96), Inches(h_px/96))
                            txBox.text_frame.text = elem.text
                        elif elem.name == 'img':
                            src_str = elem.get('src', '')
                            if "base64," in src_str:
                                image_bytes = base64.b64decode(src_str.split("base64,")[1])
                                slide.shapes.add_picture(BytesIO(image_bytes), Inches(l_px/96), Inches(t_px/96), width=Inches(w_px/96), height=Inches(h_px/96))
                    except Exception as e:
                        logging.warning(f"HTML 요소 복원 실패(스킵됨): {e}")
                        continue
            prs.save(output_pptx_path)
            return True
        except Exception as e:
            logging.error(f"HTML to PPTX 역변환 중 치명적 오류 발생: {e}")
            return False

    # =============================================================
    # 📑 [워드] DOCX ⇄ HTML (구조적 표 변환 및 안전한 조립)
    # =============================================================
    @staticmethod
    def docx_to_html(docx_path: str, output_html_path: str) -> bool:
        try:
            with open(docx_path, "rb") as docx_file:
                # mammoth는 표(Table) 구조도 표준 <table> 태그로 자동 빌드함
                result = mammoth.convert_to_html(docx_file)
            with open(output_html_path, "w", encoding="utf-8") as f:
                f.write(result.value)
            return True
        except Exception as e:
            logging.error(f"DOCX to HTML 변환 실패: {e}")
            return False

    @staticmethod
    def html_to_docx(html_path: str, output_docx_path: str) -> bool:
        try:
            doc = Document()
            with open(html_path, "r", encoding="utf-8") as f:
                soup = BeautifulSoup(f.read(), "html.parser")
                
            for elem in soup.find_all(['h1', 'h2', 'p', 'img', 'table']):
                try:
                    if elem.name == 'h1':
                        doc.add_heading(elem.text, level=1)
                    elif elem.name == 'h2':
                        doc.add_heading(elem.text, level=2)
                    elif elem.name == 'p':
                        doc.add_paragraph(elem.text)
                    elif elem.name == 'img':
                        src_str = elem.get('src', '')
                        if "base64," in src_str:
                            image_bytes = base64.b64decode(src_str.split("base64,")[1])
                            doc.add_picture(BytesIO(image_bytes))
                            
                    # ✨ 3. 표(Table) 인식 및 워드 컴포넌트 변환 추가
                    elif elem.name == 'table':
                        rows = elem.find_all('tr')
                        if not rows: continue
                        
                        # 행과 열 크기 조사 후 빈 표 생성
                        num_rows = len(rows)
                        num_cols = len(rows[0].find_all(['th', 'td']))
                        word_table = doc.add_table(rows=num_rows, cols=num_cols)
                        word_table.style = 'Table Grid' # 격자선 스타일 기본 제공
                        
                        for r_idx, row in enumerate(rows):
                            cells = row.find_all(['th', 'td'])
                            for c_idx, cell in enumerate(cells):
                                if c_idx < num_cols: # 인덱스 초과 예외 방지
                                    word_table.rows[r_idx].cells[c_idx].text = cell.text.strip()

except Exception as e:
logging.warning(f"워드 요소 복원 중 에러 발생(스킵됨): {e}")
continue
doc.save(output_docx_path)
return True
except Exception as e:
logging.error(f"HTML to DOCX 역변환 실패: {e}")
return False
```
------------------------------
## 4. 🕸️ LangGraph 에이전트 연동 (최종 수정본)
에러 발생 시 결과 값(output_path)이 아닌 에러 상태를 핸들링할 수 있도록 샌드박싱 처리를 추가한 최종 그래프 빌드 구조입니다.
```python
from typing import TypedDict, Optional
from langgraph.graph import StateGraph, END
from converter import OfflineConverter
class ConversionState(TypedDict):
input_path: str
output_path: Optional[str]
file_type: str # 'pptx' or 'docx'
direction: str # 'to_html' or 'to_office'
status: str # 'SUCCESS' or 'FAILED'
def pptx_process_node(state: ConversionState) -> dict:
in_p = state["input_path"]
if state["direction"] == "to_html":
out_p = in_p.replace(".pptx", ".html")
success = OfflineConverter.pptx_to_html(in_p, out_p)
else:
out_p = in_p.replace(".html", "_edited.pptx")
success = OfflineConverter.html_to_pptx(in_p, out_p)
return {"output_path": out_p if success else None, "status": "SUCCESS" if success else "FAILED"}
def docx_process_node(state: ConversionState) -> dict:
in_p = state["input_path"]
if state["direction"] == "to_html":
out_p = in_p.replace(".docx", ".html")
success = OfflineConverter.docx_to_html(in_p, out_p)
else:
out_p = in_p.replace(".html", "_edited.docx")
success = OfflineConverter.html_to_docx(in_p, out_p)
return {"output_path": out_p if success else None, "status": "SUCCESS" if success else "FAILED"}
def router_logic(state: ConversionState) -> str:
return "pptx" if state["file_type"] == "pptx" else "docx"
## 그래프 인스턴스화
workflow = StateGraph(ConversionState)
workflow.add_node("pptx_node", pptx_process_node)
workflow.add_node("docx_node", docx_process_node)
workflow.set_conditional_entry_point(router_logic, {"pptx": "pptx_node", "docx": "docx_node"})
workflow.add_edge("pptx_node", END)
workflow.add_edge("docx_node", END)
app = workflow.compile()
```


***

<FollowUp>
오프라인 서버에 배포하기 전에 추가로 구현하고 싶으신 기능이 있으신가요? 예를 들어:
* 프론트엔드가 보낸 JSON을 파일로 거치지 않고 **메모리(RAM) 상에서 직접 변환 처리**하기
* 표(Table) 변환 시 **배경색(정형 테마 스타일)까지 복원**하도록 고도화하기

원하시는 방향을 말씀해 주시면 코드를 더 정밀하게 다듬어 드릴게요.
</FollowUp>


