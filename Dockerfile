# 1. 베이스 이미지 (파이썬 3.11)
FROM python:3.11-slim

# 2. 환경변수 설정
# HOME=/tmp 설정 이유: LibreOffice가 실행될 때 설정 파일을 쓸 권한 문제 방지
ENV DEBIAN_FRONTEND=noninteractive \
    JAVA_HOME=/usr/lib/jvm/default-java \
    PYTHONUNBUFFERED=1 \
    HOME=/tmp

# 3. 필수 패키지 설치 (강력한 호환성 버전)
# libreoffice-writer 대신 libreoffice(전체)를 설치하여 의존성 누락 방지
# libxinerama1, libdbus-glib-1-2 등: slim 이미지에서 누락되기 쉬운 GUI 라이브러리 강제 설치
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    libreoffice \
    libreoffice-java-common \
    default-jre \
    fonts-nanum \
    fonts-noto-cjk \
    libxinerama1 \
    libfontconfig1 \
    libdbus-glib-1-2 \
    libcairo2 \
    libcups2 \
    libglu1-mesa \
    libsm6 \
    libxrender1 \
    libxt6 \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# 4. 작업 디렉토리 설정
WORKDIR /app

# 5. 라이브러리 설치
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 6. 코드 복사
COPY . .

# 7. 서버 실행
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "10000"]