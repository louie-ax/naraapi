import os
import requests
from fastapi import FastAPI, Request
from supabase import create_client
from datetime import datetime

# 환경변수 로드
SUPABASE_URL = os.environ.get("https://zxfxouwylutlxragrhzd.supabase.co")
SUPABASE_KEY = os.environ.get("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inp4ZnhvdXd5bHV0bHhyYWdyaHpkIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc2MzQ0NTc5OCwiZXhwIjoyMDc5MDIxNzk4fQ.u_Or1_1p1PU4ekYuLuzXGs1ecqqfzE1Ak9lCU05ebjU")

app = FastAPI()

# Supabase 클라이언트 연결
if SUPABASE_URL and SUPABASE_KEY:
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
else:
    supabase = None

@app.get("/")
def read_root():
    return {"status": "Worker is ready (v2.0)"}

@app.post("/parse")
async def parse_notice(request: Request):
    try:
        payload = await request.json()
        record = payload.get('record') # Webhook이 보낸 데이터

        if not record:
            return {"msg": "No record data"}
        
        # 1. 처리 상태 확인 (NEW가 아니면 스킵)
        if record.get('process_status') != 'NEW':
            print(f"SKIP: 상태가 {record.get('process_status')}입니다.")
            return {"msg": "Skipped"}

        # 2. 식별자 가져오기 (공고번호 + 차수)
        bid_no = record.get('bidNtceNo')
        bid_ord = record.get('bidNtceOrd')
        
        if not bid_no or not bid_ord:
             return {"msg": "Primary Key Missing"}

        print(f"🚀 작업 시작: 공고번호[{bid_no}]-차수[{bid_ord}]")

        # 3. 상태 변경 (PROCESSING)
        # 주의: 복합키이므로 두 가지 조건을 모두 걸어야 함 (.eq 2번 사용)
        supabase.table('bid_notices').update({'process_status': 'PROCESSING'})\
            .eq('bidNtceNo', bid_no)\
            .eq('bidNtceOrd', bid_ord).execute()

        full_text = ""
        file_count = 0

        # 4. 첨부파일 URL 1~10번 순회하며 다운로드 시뮬레이션
        for i in range(1, 11):
            url_col = f"ntceSpecDocUrl{i}"     # 컬럼명 생성: ntceSpecDocUrl1 ... 10
            name_col = f"ntceSpecFileNm{i}"    # 파일명 컬럼
            
            file_url = record.get(url_col)
            file_name = record.get(name_col, f"file_{i}")

            # URL이 있고(None 아님), 빈 문자열이 아니며, http로 시작할 때
            if file_url and str(file_url).strip() != "" and str(file_url).startswith("http"):
                print(f"📥 발견({i}): {file_name} -> {file_url}")
                
                # --- [실제 파싱 로직이 들어갈 자리] ---
                # response = requests.get(file_url)
                # text = hwp_parser(response.content)
                # ----------------------------------
                
                # 지금은 테스트용 텍스트 생성
                full_text += f"\n=== [파일 {i}: {file_name}] ===\n(내용 파싱됨...)\nURL: {file_url}\n"
                file_count += 1

        if file_count == 0:
            full_text = "첨부파일 URL이 없는 공고입니다."

        # 5. 결과 저장 (DONE) & 파싱 내용 업데이트
        supabase.table('bid_notices').update({
            'parsed_content': full_text,
            'process_status': 'DONE',
            'updated_at': datetime.now().isoformat()
        }).eq('bidNtceNo', bid_no).eq('bidNtceOrd', bid_ord).execute()

        print(f"✅ 작업 완료: {bid_no}-{bid_ord}")
        return {"status": "Success"}

    except Exception as e:
        error_msg = str(e)
        print(f"❌ 에러 발생: {error_msg}")
        
        # 에러 발생 시 DB에 로그 남기기
        if supabase and bid_no and bid_ord:
            supabase.table('bid_notices').update({
                'process_status': 'ERROR',
                'last_error': error_msg
            }).eq('bidNtceNo', bid_no).eq('bidNtceOrd', bid_ord).execute()
            
        return {"status": "Error", "msg": error_msg}