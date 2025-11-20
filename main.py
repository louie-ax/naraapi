import os
import requests
import re # 정규표현식 모듈 추가 (맨 위에 import re 추가 필수!)
import fitz  # pymupdf (일반 텍스트용)
import pdfplumber # ★ PDF 표 인식용
import olefile # HWP용
import zlib # HWP 압축 해제용
import zipfile # HWPX용
import xml.etree.ElementTree as ET # HWPX XML 파싱용
import openpyxl # Excel용
from docx import Document # DOCX용
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from io import BytesIO
from fastapi import FastAPI, Request
from supabase import create_client
from datetime import datetime
from pydantic import BaseModel
from typing import Dict, Any, Optional
import xlrd

# 환경변수 로드
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

print("\n========== [Parser Worker v9.0 XLS & Order Fix] ==========")
if SUPABASE_URL: print(f"✅ URL Loaded")
else: print("❌ URL Missing")
print("==========================================================\n")

app = FastAPI()

if SUPABASE_URL and SUPABASE_KEY:
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
else:
    supabase = None

class SupabaseWebhook(BaseModel):
    type: str = "INSERT"
    table: str = "bid_notices"
    record: Dict[str, Any]
    schema_name: str = "public"
    old_record: Optional[Dict[str, Any]] = None

# --- [v8.1 HWP 초강력 정제 및 노이즈 제거] ---
def get_hwp_text(file_bytes):
    try:
        f = BytesIO(file_bytes)
        ole = olefile.OleFileIO(f)
        dirs = ole.listdir()
        body_sections = []
        for d in dirs:
            if d[0] == "BodyText": body_sections.append(d)
        body_sections.sort(key=lambda x: int(x[1][7:]))
        
        raw_text = ""
        for section in body_sections:
            stream = ole.openstream(section)
            data = stream.read()
            try: unpacked_data = zlib.decompress(data, -15)
            except: unpacked_data = data
            decoded = unpacked_data.decode('utf-16le', errors='ignore')
            raw_text += decoded

        if "배포용 문서입니다" in raw_text or "Distribution" in raw_text:
             return "(⚠️ 암호화된 배포용 HWP 문서는 텍스트 추출이 불가능합니다)"

        # 1. 기본 정제 (한글, 영문, 숫자, 기본 기호만 허용)
        # 특수문자 범위를 조금 더 좁혀서 이상한 기호 차단
        cleaned_text = re.sub(r'[^가-힣a-zA-Z0-9\s\.\,\-\(\)\[\]\%\~\:\/]', '', raw_text)

        # 2. [핵심] 의미 없는 줄 제거 (Garbage Collection)
        # - 길이가 너무 짧은 줄 (2글자 이하)
        # - 한글이 하나도 없는 줄 (영어/숫자/기호로만 된 깨진 데이터일 확률 높음)
        final_lines = []
        for line in cleaned_text.split('\n'):
            line = line.strip()
            
            # (1) 너무 짧은 줄 제거
            if len(line) < 2: continue
            
            # (2) 'B) l', 'ppp' 같은 바이너리 잔재 제거
            # 정규식: 한글이 적어도 1글자 이상 포함되어 있거나, 
            # 숫자가 포함된 의미있는 문장(날짜, 금액 등)인 경우만 살림
            if re.search(r'[가-힣]', line) or re.search(r'[0-9]{2,}', line):
                final_lines.append(line)

        return "\n".join(final_lines) if final_lines else "(HWP 내용 없음)"

    except Exception as e: return f"(HWP 오류: {str(e)})"
    
# --- [2. HWPX 파싱] ---
def get_hwpx_text(file_bytes):
    try:
        text = ""
        with zipfile.ZipFile(BytesIO(file_bytes)) as zf:
            for name in zf.namelist():
                if name.startswith("Contents/section") and name.endswith(".xml"):
                    xml_data = zf.read(name)
                    root = ET.fromstring(xml_data)
                    for text_tag in root.iter():
                        if text_tag.tag.endswith('t') and text_tag.text:
                            text += text_tag.text + "\n"
        return text if text else "(HWPX 내용 없음)"
    except Exception as e: return f"(HWPX 오류: {str(e)})"

# --- [3-1. XLSX/XLSM (신형 엑셀) 파싱] ---
def get_xlsx_text(file_bytes):
    try:
        wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
        text = ""
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            text += f"\n### [Sheet: {sheet}]\n"
            rows = list(ws.iter_rows(values_only=True))
            if not rows: continue
            
            # 헤더
            headers = [str(cell).replace('\n', ' ').strip() if cell is not None else "" for cell in rows[0]]
            text += "| " + " | ".join(headers) + " |\n"
            text += "| " + " | ".join(["---"] * len(headers)) + " |\n"
            
            # 데이터
            for row in rows[1:]:
                row_cells = [str(cell).replace('\n', ' ').strip() if cell is not None else "" for cell in row]
                if any(row_cells):
                    text += "| " + " | ".join(row_cells) + " |\n"
            text += "\n"
        return text
    except Exception as e: return f"(XLSX 오류: {str(e)})"

# --- [3-2. XLS (구형 엑셀) 파싱 - 신규 추가] ---
def get_xls_text(file_bytes):
    try:
        # xlrd 라이브러리로 구형 엑셀 열기
        wb = xlrd.open_workbook(file_contents=file_bytes)
        text = ""
        
        for sheet_name in wb.sheet_names():
            ws = wb.sheet_by_name(sheet_name)
            text += f"\n### [Sheet: {sheet_name}]\n"
            
            if ws.nrows == 0: continue

            # 헤더 (첫 번째 행)
            headers = [str(ws.cell_value(0, c)).replace('\n', ' ').strip() for c in range(ws.ncols)]
            text += "| " + " | ".join(headers) + " |\n"
            text += "| " + " | ".join(["---"] * len(headers)) + " |\n"

            # 데이터 (두 번째 행부터)
            for r in range(1, ws.nrows):
                row_cells = [str(ws.cell_value(r, c)).replace('\n', ' ').strip() for c in range(ws.ncols)]
                # 빈 행 체크
                if any(row_cells):
                    text += "| " + " | ".join(row_cells) + " |\n"
            text += "\n"
            
        return text
    except Exception as e: return f"(XLS 오류: {str(e)})"


# --- [4. DOCX 파싱 - ★ 순서 정렬 완벽 해결 ★] ---
def iter_block_items(parent):
    """
    문서 내부의 XML 요소를 순서대로 순회하는 제너레이터 함수
    """
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        return

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def get_docx_text(file_bytes):
    try:
        doc = Document(BytesIO(file_bytes))
        full_text = []
        
        # 문단과 표를 순서대로 가져옴
        for block in iter_block_items(doc):
            
            # 1. 문단(Paragraph)인 경우
            if isinstance(block, Paragraph):
                if block.text.strip():
                    full_text.append(block.text)
            
            # 2. 표(Table)인 경우 -> 마크다운 변환
            elif isinstance(block, Table):
                if not block.rows: continue
                
                full_text.append("\n") # 표 위 공백
                rows_data = []
                
                for row in block.rows:
                    row_cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
                    rows_data.append(row_cells)
                
                if not rows_data: continue

                # 헤더
                headers = rows_data[0]
                full_text.append("| " + " | ".join(headers) + " |")
                full_text.append("| " + " | ".join(["---"] * len(headers)) + " |")
                
                # 데이터
                for row in rows_data[1:]:
                    full_text.append("| " + " | ".join(row) + " |")
                
                full_text.append("\n") # 표 아래 공백

        return "\n".join(full_text)
    except Exception as e:
        return f"(DOCX 오류: {str(e)})"

# --- [5. PDF 파싱 (기존 유지)] ---
def get_pdf_text(file_bytes):
    try:
        text_output = []
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            for page_num, page in enumerate(pdf.pages):
                text_output.append(f"\n### [Page {page_num + 1}]")
                
                tables = page.find_tables()
                table_bboxes = [table.bbox for table in tables]

                def not_within_tables(obj):
                    if obj["object_type"] == "char":
                        x0, top, x1, bottom = obj["x0"], obj["top"], obj["x1"], obj["bottom"]
                        cx, cy = (x0 + x1) / 2, (top + bottom) / 2
                        for bbox in table_bboxes:
                            if bbox[0] <= cx <= bbox[2] and bbox[1] <= cy <= bbox[3]:
                                return False
                    return True

                filtered_page = page.filter(not_within_tables)
                raw_text = filtered_page.extract_text(layout=True)
                if raw_text: text_output.append(raw_text)

                extracted_tables = page.extract_tables()
                if extracted_tables:
                    text_output.append("\n")
                    for table in extracted_tables:
                        clean_table = []
                        for row in table:
                            clean_row = [str(cell).replace('\n', ' ').strip() if cell is not None else "" for cell in row]
                            if any(clean_row): clean_table.append(clean_row)
                        
                        if not clean_table: continue
                        
                        headers = clean_table[0]
                        text_output.append("| " + " | ".join(headers) + " |")
                        text_output.append("| " + " | ".join(["---"] * len(headers)) + " |")
                        for row in clean_table[1:]:
                            text_output.append("| " + " | ".join(row) + " |")
                        text_output.append("\n")

        return "\n".join(text_output)
    except Exception as e:
        try:
            t = ""
            with fitz.open(stream=file_bytes, filetype="pdf") as doc:
                for p in doc: t += p.get_text()
            return t
        except: return f"(PDF 오류: {str(e)})"

# --- [통합 핸들러] ---
def extract_text_from_file(file_bytes, ext):
    ext = ext.lower()
    try:
        if 'pdf' in ext: return get_pdf_text(file_bytes)
        elif 'docx' in ext or 'doc' in ext: return get_docx_text(file_bytes)
        elif 'hwp' == ext: return get_hwp_text(file_bytes)
        elif 'hwpx' in ext: return get_hwpx_text(file_bytes)
        elif 'xlsx' in ext or 'xlsm' in ext: return get_xlsx_text(file_bytes)
        elif 'xls' == ext: return get_xls_text(file_bytes) # ★ xls 추가
        else: return f"(지원하지 않는 파일: {ext})"
    except Exception as e: return f"시스템 에러: {str(e)}"

# --- [엔드포인트] ---
@app.get("/")
def read_root():
    return {"status": "Worker v9.0 (XLS & DOCX Order Fix)"}

@app.post("/parse")
async def parse_notice(payload: SupabaseWebhook):
    if not supabase: return {"status": "Error", "msg": "Env Vars Missing"}

    try:
        record = payload.record
        if not record: return {"msg": "No record"}

        bid_no = record.get('bidNtceNo')
        bid_ord = record.get('bidNtceOrd')
        if record.get('process_status') != 'NEW': return {"msg": "Skipped"}
        
        print(f"🚀 [시작] {bid_no}-{bid_ord}")
        supabase.table('bid_notices').update({'process_status': 'PROCESSING'}).eq('bidNtceNo', bid_no).eq('bidNtceOrd', bid_ord).execute()
        
        full_log = ""
        for i in range(1, 11):
            url = record.get(f"ntceSpecDocUrl{i}")
            name = record.get(f"ntceSpecFileNm{i}", f"File_{i}")

            if url and str(url).startswith("http"):
                print(f"  📥 [다운로드] {name}")
                try:
                    resp = requests.get(url, timeout=60)
                    text = extract_text_from_file(resp.content, name.split('.')[-1])
                    
                    supabase.table('bid_attachments').insert({
                        "bidNtceNo": bid_no, "bidNtceOrd": bid_ord,
                        "file_name": name, "file_url": url, "file_type": name.split('.')[-1],
                        "extracted_text": text
                    }).execute()
                    full_log += f"  💾 [성공] {name} ({len(text)}자)\n"
                except Exception as e:
                    full_log += f"  ❌ [실패] {name}\n"
                    print(f"실패: {e}")

        supabase.table('bid_notices').update({'process_status': 'DONE', 'parsed_content': full_log, 'updated_at': datetime.now().isoformat()}).eq('bidNtceNo', bid_no).eq('bidNtceOrd', bid_ord).execute()
        print(f"✅ [완료] {bid_no}")
        return {"status": "Success"}

    except Exception as e:
        print(f"❌ [에러] {e}")
        return {"status": "Error", "msg": str(e)}