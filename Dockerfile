# 1. 베이스 이미지 (파이썬 3.11)
FROM python:3.11-slim

# 2. 환경변수 설정
ENV DEBIAN_FRONTEND=noninteractive \
    JAVA_HOME=/usr/lib/jvm/default-java \
    PYTHONUNBUFFERED=1

# 3. 필수 패키지 설치 (LibreOffice 및 의존성 강화)
# libreoffice-java-common: Java 의존성 해결
# fonts-nanum: 한글 폰트 깨짐 방지
# libgl1-mesa-glx, libxinerama1 등: Headless 실행 시 필요한 그래픽 라이브러리
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    libreoffice-writer \
    libreoffice-java-common \
    default-jre \
    fonts-nanum \
    libgl1-mesa-glx \
    libxinerama1 \
    libxcursor1 \
    libxrandr2 \
    libxi6 \
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