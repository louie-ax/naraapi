# 1. 베이스 이미지
FROM python:3.11-slim

# 2. 필수 패키지 설치 (Java 경로 설정 포함)
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-writer \
    fonts-nanum \
    default-jre \
    libreoffice-java-common \
    && rm -rf /var/lib/apt/lists/*

# ★ [추가] Java 환경변수 설정 (LibreOffice가 Java를 찾도록 도와줌)
ENV JAVA_HOME=/usr/lib/jvm/default-java

# 3. 작업 디렉토리
WORKDIR /app

# 4. 라이브러리 설치
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 5. 코드 복사
COPY . .

# 6. 실행
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "10000"]