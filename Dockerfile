# 1. 베이스 이미지 (파이썬 3.11)
FROM python:3.11-slim

# 2. LibreOffice 및 한글 폰트, 필수 도구 설치
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-writer \
    fonts-nanum \
    default-jre \
    && rm -rf /var/lib/apt/lists/*

# 3. 작업 디렉토리 설정
WORKDIR /app

# 4. 라이브러리 설치
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 5. 코드 복사
COPY . .

# 6. 서버 실행 (포트 10000)
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "10000"]