import os
import requests
from fastapi import FastAPI, Request
from supabase import create_client

# Railway에서 설정할 환경변수
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

app = FastAPI()
# 클라이언트 생성 (키가 없으면 에러 방지를 위해 None 처리)
if SUPABASE_URL and SUPABASE_KEY:
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
else:
    supabase = None

@app.get("/")
def read_root():
    return {"status": "Server is ALIVE!"}

@app.post("/parse")
async def parse_notice(request: Request):
    print("🔔 Webhook 신호 수신!")
    
    try:
        payload = await request.json()
        record = payload.get('record') # Supabase가 보낸 데이터

        if not record:
            return {"msg": "No record data"}
        
        # 이미 처리된 건은 패스
        if record.get('process_status') != 'NEW':
            print(f"PASS: 상태가 {record.get('process_status')}입니다.")
            return {"msg": "Skipped"}

        # 공고번호 확인
        bid_no = record.get('bidNtceNo')
        bid_ord = record.get('bidNtceOrd')
        print(f"🚀 작업 시작: {bid_no}-{bid_ord}")

        # 1. 상태 변경 (PROCESSING)
        supabase.table('bid_notices').update({'process_status': 'PROCESSING'})\
            .eq('bidNtceNo', bid_no).eq('bidNtceOrd', bid_ord).execute()

        # 2. 파싱 로직 (테스트용)
        full_text = ""
        
        # URL 1번부터 5번까지만 확인해봅니다
        for i in range(1, 6):
            url_key = f"ntceSpecDocUrl{i}"
            file_url = record.get(url_key)
            
            if file_url and str(file_url).startswith("http"):
                print(f"📥 파일 발견: {file_url}")
                # 여기서 실제 다운로드/파싱을 수행합니다.
                # 지금은 성공했다고 가정하고 URL만 기록합니다.
                full_text += f"[File {i} Processed: {file_url}]\n"

        # 3. 결과 저장 (DONE)
        supabase.table('bid_notices').update({
            'parsed_content': full_text,
            'process_status': 'DONE',
            'updated_at': 'now()'
        }).eq('bidNtceNo', bid_no).eq('bidNtceOrd', bid_ord).execute()

        print(f"✅ 작업 완료: {bid_no}")
        return {"status": "Success"}

    except Exception as e:
        print(f"❌ 에러 발생: {str(e)}")
        return {"status": "Error", "msg": str(e)}