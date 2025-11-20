import os
import requests
import fitz  # pymupdf (일반 텍스트용)
import pdfplumber # ★ PDF 표 인식용
import olefile # HWP용
import zlib # HWP 압축 해제용
import zipfile # HWPX용
import xml.etree.ElementTree as ET # HWPX XML 파싱용
import openpyxl # Excel용
from docx import Document # DOCX용
from io import BytesIO
from fastapi import FastAPI, Request
from supabase import create_client
from datetime import datetime
from pydantic import BaseModel
from typing import Dict, Any, Optional

# 환경변수 로드
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

print("\n========== [Parser Worker v6.0 Table Master] ==========")
if SUPABASE_URL: print(f"✅ URL Loaded")
else: print("❌ URL Missing")
print("=======================================================\n")

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

# --- [1. HWP (OLE) 파싱 - 텍스트 순차 추출] ---
# --- [수정된 HWP 파싱 함수 (v7.0)] ---
def get_hwp_text(file_bytes):
    try:
        f = BytesIO(file_bytes)
        ole = olefile.OleFileIO(f)
        
        dirs = ole.listdir()
        body_sections = []
        
        # BodyText 섹션 찾기
        for d in dirs:
            if d[0] == "BodyText":
                body_sections.append(d)
        
        # 섹션 순서대로 정렬
        body_sections.sort(key=lambda x: int(x[1][7:]))
        
        text = ""
        for section in body_sections:
            stream = ole.openstream(section)
            data = stream.read()
            
            # 압축 해제 시도 (HWP 5.0+)
            try:
                unpacked_data = zlib.decompress(data, -15)
            except:
                unpacked_data = data
            
            # UTF-16LE 디코딩
            decoded = unpacked_data.decode('utf-16le', errors='ignore')
            
            # --- [★ 핵심 수정: 강력한 텍스트 필터링 ★] ---
            # 한글(가-힣), 영문, 숫자, 기본 특수문자, 공백만 허용
            # 나머지 제어문자나 외계어는 모두 제거
            filtered_text = ""
            for char in decoded:
                # 한글 범위: 0xAC00 ~ 0xD7A3
                # 영문/숫자/특수문자: 0x0020 ~ 0x007E
                # 줄바꿈/탭: \n, \t, \r
                code = ord(char)
                if (0xAC00 <= code <= 0xD7A3) or \
                   (0x0020 <= code <= 0x007E) or \
                   char in ['\n', '\t', '\r', ' ']:
                    filtered_text += char
            
            text += filtered_text + "\n"
            
        # 결과가 비어있거나 너무 짧으면(헤더만 읽은 경우) 실패 처리
        if len(text.strip()) < 10:
             return "(HWP 텍스트 추출 실패 - 내용 없음)"

        return text
        
    except Exception as e:
        return f"(HWP 파싱 시스템 에러: {str(e)})"

# --- [2. HWPX (XML) 파싱] ---
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
    except Exception as e:
        return f"(HWPX 파싱 실패: {str(e)})"

# --- [3. Excel 파싱 - 마크다운 표 변환] ---
def get_excel_text(file_bytes):
    try:
        wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
        text = ""
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            text += f"\n### 시트명: {sheet}\n"
            
            # 엑셀 내용을 마크다운 표로 변환
            rows = list(ws.iter_rows(values_only=True))
            if not rows: continue

            # 헤더 생성 (첫 줄)
            headers = rows[0]
            header_str = "| " + " | ".join([str(h) if h else " " for h in headers]) + " |"
            separator = "| " + " | ".join(["---"] * len(headers)) + " |"
            
            text += header_str + "\n" + separator + "\n"
            
            # 데이터 생성
            for row in rows[1:]:
                row_str = "| " + " | ".join([str(cell).replace("\n", " ") if cell is not None else " " for cell in row]) + " |"
                text += row_str + "\n"
            text += "\n"

        return text
    except Exception as e:
        return f"(Excel 파싱 실패: {str(e)})"

# --- [4. DOCX 파싱 - ★ 표 인식 강화 ★] ---
def get_docx_text(file_bytes):
    try:
        doc = Document(BytesIO(file_bytes))
        full_text = []
        
        # 문서의 요소들을 순서대로 읽는 것은 python-docx에서 어렵습니다.
        # 대신, 문단을 먼저 다 읽고 -> 그 다음 표를 마크다운으로 변환해서 붙입니다.
        
        # 1. 일반 텍스트 (Paragraphs)
        full_text.append("=== [문서 본문] ===")
        for para in doc.paragraphs:
            if para.text.strip():
                full_text.append(para.text)
        
        # 2. 표 (Tables) -> 마크다운 변환
        if doc.tables:
            full_text.append("\n=== [표 데이터 (Table)] ===")
            for table in doc.tables:
                # 표가 비어있으면 패스
                if not table.rows: continue
                
                # 각 행을 처리
                for i, row in enumerate(table.rows):
                    # 셀 안의 줄바꿈은 공백으로 치환 (표 깨짐 방지)
                    row_cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
                    row_str = "| " + " | ".join(row_cells) + " |"
                    full_text.append(row_str)
                    
                    # 첫 번째 행(헤더) 밑에 구분선 추가
                    if i == 0:
                        sep_str = "| " + " | ".join(["---"] * len(row_cells)) + " |"
                        full_text.append(sep_str)
                
                full_text.append("") # 표 사이 공백
                
        return "\n".join(full_text)
    except Exception as e:
        return f"(DOCX 파싱 실패: {str(e)})"

# --- [5. PDF 파싱 - ★ pdfplumber 도입 ★] ---
def get_pdf_text(file_bytes):
    try:
        text = ""
        # pdfplumber로 열기
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                # 1. 텍스트 추출
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
                
                # 2. 표 추출 (마크다운 변환)
                tables = page.extract_tables()
                if tables:
                    text += "\n[PDF 표 데이터]\n"
                    for table in tables:
                        # table은 리스트의 리스트 형태 [[값, 값], [값, 값]]
                        # None 값 제거 및 줄바꿈 제거
                        clean_table = [[(str(cell).replace("\n", " ") if cell else "") for cell in row] for row in table]
                        
                        if not clean_table: continue

                        # 헤더 처리
                        headers = clean_table[0]
                        header_str = "| " + " | ".join(headers) + " |"
                        sep_str = "| " + " | ".join(["---"] * len(headers)) + " |"
                        
                        text += header_str + "\n" + sep_str + "\n"
                        
                        # 데이터 처리
                        for row in clean_table[1:]:
                            row_str = "| " + " | ".join(row) + " |"
                            text += row_str + "\n"
                        text += "\n"
                        
        return text
    except Exception as e:
        # pdfplumber 실패 시 fitz(pymupdf)로 백업 시도
        try:
            text = ""
            with fitz.open(stream=file_bytes, filetype="pdf") as doc:
                for page in doc: text += page.get_text()
            return text
        except Exception as e2:
            return f"(PDF 파싱 실패: {str(e)})"

# --- [통합 파싱 핸들러] ---
def extract_text_from_file(file_bytes, ext):
    ext = ext.lower()
    try:
        if 'pdf' in ext:
            return get_pdf_text(file_bytes)
        elif 'docx' in ext or 'doc' in ext:
            return get_docx_text(file_bytes)
        elif 'hwp' == ext:
            return get_hwp_text(file_bytes)
        elif 'hwpx' in ext:
            return get_hwpx_text(file_bytes)
        elif 'xlsx' in ext or 'xlsm' in ext:
            return get_excel_text(file_bytes)
        else:
            return f"(지원하지 않는 파일 형식: {ext})"
    except Exception as e:
        return f"시스템 처리 에러: {str(e)}"

# -----------------------

@app.get("/")
def read_root():
    return {"status": "Worker Ready (v6.0 Table Support)"}

@app.post("/parse")
async def parse_notice(payload: SupabaseWebhook):
    if not supabase: return {"status": "Error", "msg": "Env Vars Missing"}

    try:
        record = payload.record
        if not record: return {"msg": "No record"}

        bid_no = record.get('bidNtceNo')
        bid_ord = record.get('bidNtceOrd')
        status = record.get('process_status')
        
        if status != 'NEW':
            print(f"⛔ [스킵] {bid_no} 상태: {status}")
            return {"msg": "Skipped"}
        
        print(f"🚀 [시작] {bid_no}-{bid_ord}")

        supabase.table('bid_notices').update({'process_status': 'PROCESSING'})\
            .eq('bidNtceNo', bid_no).eq('bidNtceOrd', bid_ord).execute()

        full_report = ""

        for i in range(1, 11):
            url_col = f"ntceSpecDocUrl{i}"
            name_col = f"ntceSpecFileNm{i}"
            
            file_url = record.get(url_col)
            file_name = record.get(name_col, f"File_{i}")

            if file_url and str(file_url).startswith("http"):
                print(f"  📥 [다운로드] {file_name}")
                try:
                    response = requests.get(file_url, timeout=60)
                    file_bytes = response.content
                    ext = file_name.split('.')[-1].lower()
                    
                    parsed_text = extract_text_from_file(file_bytes, ext)
                    
                    supabase.table('bid_attachments').insert({
                        "bidNtceNo": bid_no,
                        "bidNtceOrd": bid_ord,
                        "file_name": file_name,
                        "file_url": file_url,
                        "file_type": ext,
                        "extracted_text": parsed_text
                    }).execute()
                    
                    log = f"  💾 [성공] {file_name} ({len(parsed_text)}자)"
                    print(log)
                    full_report += log + "\n"
                    
                except Exception as e:
                    print(f"  ⚠️ [실패] {file_name}: {e}")
                    full_report += f"  ❌ [에러] {file_name}\n"

        supabase.table('bid_notices').update({
            'process_status': 'DONE',
            'parsed_content': full_report,
            'updated_at': datetime.now().isoformat()
        }).eq('bidNtceNo', bid_no).eq('bidNtceOrd', bid_ord).execute()

        print(f"✅ [종료] {bid_no}")
        return {"status": "Success"}

    except Exception as e:
        print(f"❌ [시스템 에러] {e}")
        return {"status": "Error", "msg": str(e)}