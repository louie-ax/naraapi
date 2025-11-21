# 1. 베이스 이미지
FROM python:3.11-slim

# 2. 필수 패키지 및 LibreOffice 의존성 설치 (중요!)
# libreoffice-java-common과 함께 headless 실행에 필요한 그래픽 라이브러리들을 추가합니다.
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-writer \
    fonts-nanum \
    default-jre \
    libreoffice-java-common \
    libxinerama1 libx11-xcb1 libxrandr2 libxi6 libgl1-mesa-glx \
    && rm -rf /var/lib/apt/lists/*

# 3. Java 환경변수 설정
ENV JAVA_HOME=/usr/lib/jvm/default-java

# 4. 작업 디렉토리 설정
WORKDIR /app

# 5. 라이브러리 설치
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 6. 코드 복사
COPY . .

# 7. 서버 실행
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "10000"]