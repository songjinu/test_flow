# 📦 오프라인 통합 문서 변환 시스템 가이드 (PPTX / DOCX ⇄ HTML)

인터넷 연결이 불가능한 폐쇄망 환경에서 대형 AI 모델(Docling 등)이나 유료 라이브러리(Aspose) 없이, **순수 파이썬 초경량 패키지만을 활용하여 이미지와 좌표를 보존하는 양방향 변환 시스템** 가이드입니다.

---

## 1. ⚙️ 오프라인 환경 구축 (패키지 사전 준비)

인터넷이 되는 외부 환경에서 필요한 필수 패키지(`.whl`)를 미리 다운로드하여 오프라인 서버로 반입해야 합니다.

### 1) 외부망 PC에서 패키지 다운로드
```bash
mkdir offline_packages
cd offline_packages

# 필수 패키지 3종 한 번에 다운로드
pip download langgraph python-pptx python-docx mammoth beautifulsoup4
```

### 2) 폐쇄망 서버에서 패키지 설치
다운로드한 폴더를 오프라인 서버로 이동한 후 로컬 경로 지정을 통해 설치합니다.
```bash
pip install --no-index --find-links=/path/to/offline_packages langgraph python-pptx python-docx mammoth beautifulsoup4
```

---

## 2. 🛠️ 통합 양방향 변환 모듈 코드 (`converter.py`)

이미지를 **Base64 인라인 텍스트(`data:image/...;base64`)**로 변환하여 인터넷이 없는 환경에서도 단 하나의 HTML 파일 안에 모든 텍스트와 그림이 정상 출력되도록 설계된 통합 소스코드입니다.

```python
import os
import re
import base64
import mammoth
from io import BytesIO
from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches
from bs4 import BeautifulSoup

class OfflineConverter:
    """인터넷 없이 로컬에서 구동되는 이미지 및 좌표 포함 오피스 양방향 변환 모듈"""
    
    # =============================================================
    # 📊 [파워포인트] PPTX ⇄ HTML (절대좌표 및 이미지 원본 추출)
    # =============================================================
    @staticmethod
    def pptx_to_html(pptx_path: str, output_html_path: str):
        prs = Presentation(pptx_path)
        sw, sh = int(prs.slide_width / 9525), int(prs.slide_height / 9525)
        
        html = f"<html><head><style>.slide {{ width:{sw}px; height:{sh}px; position:relative; background:white; border:1px solid #ccc; margin-bottom:20px; }} .element {{ position:absolute; }}</style></head><body>"
        
        for i, slide in enumerate(prs.slides):
            html += f"<div class='slide' id='slide-{i+1}'>"
            for shape in slide.shapes:
                l, t = int(shape.left / 9525), int(shape.top / 9525)
                w, h = int(shape.width / 9525), int(shape.height / 9525)
                style = f"left:{l}px; top:{t}px; width:{w}px; height:{h}px;"
                
                # 텍스트 박스 추출
                if hasattr(shape, "text") and shape.text.strip() and shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                    html += f"<div class='element' style='{style}'>{shape.text}</div>"
                
                # 이미지 추출 (로컬에서 즉시 Base64 텍스트 변환)
                elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    base64_str = base64.b64encode(shape.image.blob).decode('utf-8')
                    img_src = f"data:{shape.image.content_type};base64,{base64_str}"
                    html += f"<img class='element' style='{style}' src='{img_src}' />"
            html += "</div>"
        html += "</body></html>"
        
        with open(output_html_path, "w", encoding="utf-8") as f:
            f.write(html)

    @staticmethod
    def html_to_pptx(html_path: str, output_pptx_path: str):
        prs = Presentation()
        blank_layout = prs.slide_layouts[6] # 공백 레이아웃 지정
        
        with open(html_path, "r", encoding="utf-8") as f:
            soup = BeautifulSoup(f.read(), "html.parser")
            
        for s in soup.find_all(class_='slide'):
            slide = prs.slides.add_slide(blank_layout)
            for elem in s.find_all(class_='element'):
                style_str = elem.get('style', '')
                
                # HTML 내 인라인 CSS 좌표 획득
                l_px = int(re.search(r'left:\s*(\d+)px', style_str).group(1))
                t_px = int(re.search(r'top:\s*(\d+)px', style_str).group(1))
                w_px = int(re.search(r'width:\s*(\d+)px', style_str).group(1))
                h_px = int(re.search(r'height:\s*(\d+)px', style_str).group(1))
                
                # 텍스트 컴포넌트 복원
                if elem.name == 'div':
                    txBox = slide.shapes.add_textbox(Inches(l_px/96), Inches(t_px/96), Inches(w_px/96), Inches(h_px/96))
                    txBox.text_frame.text = elem.text
                # 이미지 컴포넌트 복원
                elif elem.name == 'img':
                    src_str = elem.get('src', '')
                    if "base64," in src_str:
                        image_bytes = base64.b64decode(src_str.split("base64,")[1])
                        slide.shapes.add_picture(BytesIO(image_bytes), Inches(l_px/96), Inches(t_px/96), width=Inches(w_px/96), height=Inches(h_px/96))
        prs.save(output_pptx_path)

    # =============================================================
    # 📑 [워드] DOCX ⇄ HTML (구조적 문서 매핑 및 이미지 인라인화)
    # =============================================================
    @staticmethod
    def docx_to_html(docx_path: str, output_html_path: str):
        with open(docx_path, "rb") as docx_file:
            # mammoth 라이브러리는 내장 엔진이 이미지를 Base64화하여 HTML에 매핑함
            result = mammoth.convert_to_html(docx_file)
        with open(output_html_path, "w", encoding="utf-8") as f:
            f.write(result.value)

    @staticmethod
    def html_to_docx(html_path: str, output_docx_path: str):
        doc = Document()
        with open(html_path, "r", encoding="utf-8") as f:
            soup = BeautifulSoup(f.read(), "html.parser")
            
        for elem in soup.find_all(['h1', 'h2', 'p', 'img']):
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
        doc.save(output_docx_path)
```

---

## 3. 🕸️ LangGraph 에이전트 연동 예시

생성한 `OfflineConverter` 모듈을 LangGraph 워크플로우의 실행 **Node(노드)**로 결합하여 파이프라인화하는 방법입니다. 외부 LLM 의존성 없이 입력 파라미터 값에 따라 분기 연산을 수행합니다.

```python
from typing import TypedDict, Optional
from langgraph.graph import StateGraph, END
from converter import OfflineConverter

# 1. 상태 객체 정의
class ConversionState(TypedDict):
    input_path: str
    output_path: Optional[str]
    file_type: str    # 'pptx' 또는 'docx'
    direction: str    # 'to_html' 또는 'to_office'

# 2. PPTX 노드 기능 정의
def pptx_process_node(state: ConversionState) -> dict:
    in_p = state["input_path"]
    if state["direction"] == "to_html":
        out_p = in_p.replace(".pptx", ".html")
        OfflineConverter.pptx_to_html(in_p, out_p)
    else:
        out_p = in_p.replace(".html", "_edited.pptx")
        OfflineConverter.html_to_pptx(in_p, out_p)
    return {"output_path": out_p}

# 3. DOCX 노드 기능 정의
def docx_process_node(state: ConversionState) -> dict:
    in_p = state["input_path"]
    if state["direction"] == "to_html":
        out_p = in_p.replace(".docx", ".html")
        OfflineConverter.docx_to_html(in_p, out_p)
    else:
        out_p = in_p.replace(".html", "_edited.docx")
        OfflineConverter.html_to_docx(in_p, out_p)
    return {"output_path": out_p}

# 4. 파일 종류에 따른 결정론적 라우터 조건문
def router_logic(state: ConversionState) -> str:
    return "pptx" if state["file_type"] == "pptx" else "docx"

# 5. LangGraph 아키텍처 빌드
workflow = StateGraph(ConversionState)
workflow.add_node("pptx_node", pptx_process_node)
workflow.add_node("docx_node", docx_process_node)

workflow.set_conditional_entry_point(
    router_logic,
    {"pptx": "pptx_node", "docx": "docx_node"}
)
workflow.add_edge("pptx_node", END)
workflow.add_edge("docx_node", END)

app = workflow.compile()

# 6. 로컬 파이프라인 구동 예시 (인터넷 차단 상태에서 작동)
if __name__ == "__main__":
    inputs = {
        "input_path": "sample.pptx",
        "file_type": "pptx",
        "direction": "to_html"
    }
    result = app.invoke(inputs)
    print(f"🎉 변환 완료 파일 저장소: {result['output_path']}")
```
