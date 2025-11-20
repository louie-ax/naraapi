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
# --- [v8.0 초강력 HWP 정제 함수] ---
def get_hwp_text(file_bytes):
    try:
        f = BytesIO(file_bytes)
        ole = olefile.OleFileIO(f)
        
        dirs = ole.listdir()
        body_sections = []
        for d in dirs:
            if d[0] == "BodyText":
                body_sections.append(d)
        
        body_sections.sort(key=lambda x: int(x[1][7:]))
        
        raw_text = ""
        for section in body_sections:
            stream = ole.openstream(section)
            data = stream.read()
            try:
                unpacked_data = zlib.decompress(data, -15)
            except:
                unpacked_data = data
            
            # 1차 디코딩
            decoded = unpacked_data.decode('utf-16le', errors='ignore')
            raw_text += decoded

        # --- [★ 핵심 수정: 정규표현식 기반 강력 필터링 ★] ---
        
        # 1. 배포용 문서 체크
        if "배포용 문서입니다" in raw_text or "Distribution" in raw_text:
             return "(⚠️ 암호화된 배포용 HWP 문서는 텍스트 추출이 불가능합니다)"

        # 2. 허용할 문자만 남기기 (화이트리스트)
        # 가-힣: 한글
        # a-zA-Z0-9: 영문/숫자
        # \s: 공백/줄바꿈
        # .:(),%~- : 공고문에 자주 쓰이는 특수문자들
        # [^...] : 저 안에 없는 건 다 지워라!
        
        cleaned_text = re.sub(r'[^가-힣a-zA-Z0-9\s\.\:\(\)\,\%\~\-\“\”\’\‘\/\_]', '', raw_text)
        
        # 3. 다중 공백/줄바꿈 정리
        # 연속된 줄바꿈(\n\n\n) -> \n
        cleaned_text = re.sub(r'\n+', '\n', cleaned_text)
        # 연속된 공백(    ) -> 공백 하나
        cleaned_text = re.sub(r' +', ' ', cleaned_text)

        return cleaned_text.strip() if cleaned_text.strip() else "(HWP 텍스트 추출 실패 - 내용 없음)"

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
# --- [v8.0 PDF 파싱 함수 - 표 내용 중복 제거 및 정밀 추출] ---
def get_pdf_text(file_bytes):
    try:
        text_output = []
        
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            for page_num, page in enumerate(pdf.pages):
                text_output.append(f"\n### [Page {page_num + 1}]")

                # 1. 표(Table) 영역 찾기
                tables = page.find_tables()
                table_bboxes = [table.bbox for table in tables] # 표의 좌표(영역) 저장
                
                # 2. 텍스트 추출 (표 영역 제외)
                # filter 함수를 써서 표 영역 안에 있는 글자는 무시합니다.
                def not_within_tables(obj):
                    # obj가 글자(char)일 때만 좌표 체크
                    if obj["object_type"] == "char":
                        x0, top, x1, bottom = obj["x0"], obj["top"], obj["x1"], obj["bottom"]
                        # 글자의 중심점이 표 영역 안에 들어가는지 확인
                        cx, cy = (x0 + x1) / 2, (top + bottom) / 2
                        for bbox in table_bboxes:
                            # bbox: (x0, top, x1, bottom)
                            if bbox[0] <= cx <= bbox[2] and bbox[1] <= cy <= bbox[3]:
                                return False # 표 안에 있으면 제외
                    return True

                # 필터링된 페이지에서 텍스트 추출
                filtered_page = page.filter(not_within_tables)
                raw_text = filtered_page.extract_text(layout=True)
                
                if raw_text:
                    text_output.append(raw_text)

                # 3. 표(Table) 추출 및 마크다운 변환
                extracted_tables = page.extract_tables()
                
                if extracted_tables:
                    text_output.append("\n--- [PDF 표 데이터] ---")
                    
                    for table in extracted_tables:
                        # 데이터 정제
                        clean_table = []
                        for row in table:
                            # None -> 빈 문자열, 줄바꿈 -> 공백 치환
                            clean_row = [
                                str(cell).replace('\n', ' ').strip() if cell is not None else "" 
                                for cell in row
                            ]
                            # 내용이 있는 행만 추가
                            if any(clean_row):
                                clean_table.append(clean_row)
                        
                        if not clean_table: continue

                        # 마크다운 표 생성
                        try:
                            # (1) 헤더 (첫 번째 줄)
                            headers = clean_table[0]
                            header_str = "| " + " | ".join(headers) + " |"
                            separator = "| " + " | ".join(["---"] * len(headers)) + " |"
                            
                            text_output.append(header_str)
                            text_output.append(separator)
                            
                            # (2) 데이터 (두 번째 줄부터)
                            for row in clean_table[1:]:
                                row_str = "| " + " | ".join(row) + " |"
                                text_output.append(row_str)
                            
                            text_output.append("") # 표 사이 공백
                        except Exception as e:
                            text_output.append(f"(표 변환 중 오류: {e})")

        return "\n".join(text_output)

    except Exception as e:
        # pdfplumber 실패 시 fitz(pymupdf)로 백업
        try:
            fallback_text = ""
            with fitz.open(stream=file_bytes, filetype="pdf") as doc:
                for page in doc: fallback_text += page.get_text()
            return f"(표 인식 실패, 일반 텍스트 추출됨)\n{fallback_text}"
        except:
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