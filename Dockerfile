FROM python:3.12-slim

ENV TZ=Asia/Seoul \
    PYTHONUNBUFFERED=1

WORKDIR /app

# 의존성만 먼저 설치 (코드는 docker-compose에서 바인드 마운트)
COPY requirements.txt /tmp/requirements.txt
RUN pip install --no-cache-dir -r /tmp/requirements.txt gunicorn==23.0.0

EXPOSE 5000

# 주의: 업로드 목록이 프로세스 메모리에 캐시되므로 workers는 반드시 1로 유지.
#       동시 처리는 threads로 확보 (app.py 자체가 threaded 설계).
# timeout 600: 패키지 일괄 검증 등 장시간 요청 대비
# graceful-timeout: 종료 시 atexit의 _state.json flush 시간 확보
CMD ["gunicorn", "--bind", "0.0.0.0:5000", \
     "--worker-class", "gthread", "--workers", "1", "--threads", "16", \
     "--timeout", "600", "--graceful-timeout", "30", \
     "--access-logfile", "-", "--error-logfile", "-", \
     "app:app"]
