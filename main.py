import os
import requests
import fitz  # pymupdf (PDF용)
import olefile # HWP용
import zlib # HWP 압축 해제용
from docx import Document # DOCX용
from io import BytesIO
from fastapi import FastAPI, Request
from supabase import create_client
from datetime import datetime

# 1. 환경변수 가져오기 & 디버깅 (진실의 방)
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

print("\n========== [환경변수 로딩 점검] ==========")
if SUPABASE_URL:
    print(f"✅ URL 감지됨: {SUPABASE_URL[:15]}...")
else:
    print("❌ URL이 없습니다! (None)")

if SUPABASE_KEY:
    print(f"✅ KEY 감지됨: {SUPABASE_KEY[:10]}...")
else:
    print("❌ KEY가 없습니다! (None)")
print("==========================================\n")

app = FastAPI()

# 클라이언트 생성 (변수가 없으면 생성 안 함)
if SUPABASE_URL and SUPABASE_KEY:
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
else:
    supabase = None

# --- [HWP 파싱 함수] ---
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
            clean_text = "".join([c for c in decoded if c.isprintable() or c in ['\n', '\t', ' ']])
            text += clean_text + "\n"
            
        return text
    except Exception as e:
        return f"(HWP 파싱 실패: {str(e)})"

# --- [통합 파싱 함수 (PDF/DOCX/HWP)] ---
def extract_text_from_file(file_bytes, ext):
    text = ""
    try:
        if 'pdf' in ext:
            with fitz.open(stream=file_bytes, filetype="pdf") as doc:
                for page in doc:
                    text += page.get_text()
        elif 'docx' in ext or 'doc' in ext:
            doc = Document(BytesIO(file_bytes))
            for para in doc.paragraphs:
                text += para.text + "\n"
        elif 'hwp' in ext:
            text = get_hwp_text(file_bytes)
        else:
            text = "(지원하지 않는 파일 형식입니다)"
    except Exception as e:
        text = f"에러: {str(e)}"
    return text

# -----------------------

@app.get("/")
def read_root():
    status = "Normal" if supabase else "Error (No Env Vars)"
    return {"status": f"Worker Ready (v3.0 HWP) - {status}"}

@app.post("/parse")
async def parse_notice(request: Request):
    # 안전장치: 환경변수가 없으면 작업 거부
    if not supabase:
        return {"status": "Error", "msg": "서버 환경변수 미설정"}

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

        full_log = ""

        # URL 1~10번 순회
        for i in range(1, 11):
            url_col = f"ntceSpecDocUrl{i}"
            name_col = f"ntceSpecFileNm{i}"
            
            file_url = record.get(url_col)
            file_name = record.get(name_col, f"File_{i}")

            if file_url and str(file_url).startswith("http"):
                print(f"  📥 [다운로드] {file_name}")
                try:
                    # 1. 다운로드
                    response = requests.get(file_url, timeout=60)
                    file_bytes = response.content
                    ext = file_name.split('.')[-1].lower()
                    
                    # 2. 파싱
                    parsed_text = extract_text_from_file(file_bytes, ext)
                    
                    # 3. ★ 자식 테이블 저장 ★
                    supabase.table('bid_attachments').insert({
                        "bidNtceNo": bid_no,
                        "bidNtceOrd": bid_ord,
                        "file_name": file_name,
                        "file_url": file_url,
                        "file_type": ext,
                        "extracted_text": parsed_text
                    }).execute()
                    
                    log = f"  💾 [저장] {file_name} ({len(parsed_text)}자)"
                    print(log)
                    full_log += log + "\n"

                except Exception as e:
                    print(f"  ⚠️ [실패] {file_name}: {e}")

        # 최종 완료 처리
        supabase.table('bid_notices').update({
            'process_status': 'DONE',
            'parsed_content': full_log, # 부모에는 로그만 기록
            'updated_at': datetime.now().isoformat()
        }).eq('bidNtceNo', bid_no).eq('bidNtceOrd', bid_ord).execute()

        print(f"✅ [종료] {bid_no}")
        return {"status": "Success"}

    except Exception as e:
        print(f"❌ [시스템 에러] {e}")
        return {"status": "Error", "msg": str(e)}