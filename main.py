import os
import requests
import fitz  # pymupdf (PDF용)
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

# 1. 환경변수 로드
SUPABASE_URL = os.environ.get("https://zxfxouwylutlxragrhzd.supabase.co")
SUPABASE_KEY = os.environ.get("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inp4ZnhvdXd5bHV0bHhyYWdyaHpkIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc2MzQ0NTc5OCwiZXhwIjoyMDc5MDIxNzk4fQ.u_Or1_1p1PU4ekYuLuzXGs1ecqqfzE1Ak9lCU05ebjU")

print("\n========== [Parser Worker v4.0] ==========")
if SUPABASE_URL: print(f"✅ URL: {SUPABASE_URL[:15]}...")
else: print("❌ URL Missing")
if SUPABASE_KEY: print(f"✅ KEY: {SUPABASE_KEY[:10]}...")
else: print("❌ KEY Missing")
print("==========================================\n")

app = FastAPI()

if SUPABASE_URL and SUPABASE_KEY:
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
else:
    supabase = None

# --- [1. HWP (OLE) 파싱 함수] ---
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
        
        text = ""
        for section in body_sections:
            stream = ole.openstream(section)
            data = stream.read()
            try:
                unpacked_data = zlib.decompress(data, -15)
            except:
                unpacked_data = data
            
            decoded = unpacked_data.decode('utf-16le', errors='ignore')
            
            # 배포용 문서 체크 (암호화된 경우)
            if "배포용 문서입니다" in decoded or "Distribution" in decoded:
                return "(⚠️ 암호화된 배포용 HWP 문서는 텍스트 추출이 불가능합니다)"

            clean_text = "".join([c for c in decoded if c.isprintable() or c in ['\n', '\t', ' ']])
            text += clean_text + "\n"
            
        return text
    except Exception as e:
        return f"(HWP 파싱 실패: {str(e)})"

# --- [2. HWPX (XML) 파싱 함수] ---
def get_hwpx_text(file_bytes):
    try:
        text = ""
        with zipfile.ZipFile(BytesIO(file_bytes)) as zf:
            # HWPX는 zip 안에 xml 파일들이 들어있는 구조입니다.
            for name in zf.namelist():
                if name.startswith("Contents/section") and name.endswith(".xml"):
                    xml_data = zf.read(name)
                    root = ET.fromstring(xml_data)
                    
                    # 네임스페이스 처리 및 텍스트 태그(<hp:t>) 추출
                    # HWPX 구조상 텍스트는 <hp:t> 태그 안에 있습니다.
                    for text_tag in root.iter():
                        if text_tag.tag.endswith('t'): # <hp:t>
                            if text_tag.text:
                                text += text_tag.text + "\n"
        return text if text else "(HWPX 내용 없음)"
    except Exception as e:
        return f"(HWPX 파싱 실패: {str(e)})"

# --- [3. Excel (XLSX/XLSM) 파싱 함수] ---
def get_excel_text(file_bytes):
    try:
        wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
        text = ""
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            text += f"\n--- Sheet: {sheet} ---\n"
            for row in ws.iter_rows(values_only=True):
                # None 값 제외하고 텍스트로 변환하여 합침
                row_text = " | ".join([str(cell) for cell in row if cell is not None])
                if row_text.strip():
                    text += row_text + "\n"
        return text
    except Exception as e:
        return f"(Excel 파싱 실패: {str(e)})"

# --- [4. DOCX (Word) 파싱 함수 - 테이블 포함] ---
def get_docx_text(file_bytes):
    try:
        doc = Document(BytesIO(file_bytes))
        text = ""
        
        # 1. 문단(Paragraphs) 추출
        for para in doc.paragraphs:
            text += para.text + "\n"
            
        # 2. 표(Tables) 추출 (순서가 뒤섞일 수 있으나 내용은 확보됨)
        if doc.tables:
            text += "\n[표 내용 추출]\n"
            for table in doc.tables:
                for row in table.rows:
                    row_text = " | ".join([cell.text.strip() for cell in row.cells])
                    text += row_text + "\n"
                text += "\n"
        return text
    except Exception as e:
        return f"(DOCX 파싱 실패: {str(e)})"

# --- [통합 파싱 핸들러] ---
def extract_text_from_file(file_bytes, ext):
    ext = ext.lower()
    text = ""
    try:
        # PDF
        if 'pdf' in ext:
            with fitz.open(stream=file_bytes, filetype="pdf") as doc:
                for page in doc:
                    text += page.get_text() # 기본 텍스트 추출
        
        # Word
        elif 'docx' in ext or 'doc' in ext:
            text = get_docx_text(file_bytes)
            
        # HWP (Legacy)
        elif 'hwp' == ext: # hwp (not hwpx)
            text = get_hwp_text(file_bytes)
            
        # HWPX (XML based)
        elif 'hwpx' in ext:
            text = get_hwpx_text(file_bytes)
            
        # Excel
        elif 'xlsx' in ext or 'xlsm' in ext:
            text = get_excel_text(file_bytes)
            
        else:
            text = "(지원하지 않는 파일 형식입니다)"
    except Exception as e:
        text = f"시스템 처리 에러: {str(e)}"
    return text

# -----------------------

@app.get("/")
def read_root():
    status = "Normal" if supabase else "Error (No Env Vars)"
    return {"status": f"Worker Ready (v4.0 All-in-One) - {status}"}

@app.post("/parse")
async def parse_notice(request: Request):
    if not supabase: return {"status": "Error", "msg": "Env Vars Missing"}

    try:
        payload = await request.json()
        record = payload.get('record')

        if not record or record.get('process_status') != 'NEW':
            return {"msg": "Skipped"}

        bid_no = record.get('bidNtceNo')
        bid_ord = record.get('bidNtceOrd')
        
        print(f"🚀 [시작] {bid_no}-{bid_ord}")

        # 상태 변경 (PROCESSING)
        supabase.table('bid_notices').update({'process_status': 'PROCESSING'})\
            .eq('bidNtceNo', bid_no).eq('bidNtceOrd', bid_ord).execute()

        full_report = ""

        # URL 1~10번 순회
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
                    
                    # 파싱 수행
                    parsed_text = extract_text_from_file(file_bytes, ext)
                    
                    # 자식 테이블 저장
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

        # 최종 완료
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