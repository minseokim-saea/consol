// 연결 재무보고 통합 시스템 — 배포 파이프라인
// 전제: Jenkins가 운영 서버(우분투)에서 직접 실행되고, jenkins 계정이
//       docker 그룹 소속 + /srv/consol-app 쓰기 권한을 가짐 (배포 가이드 7장 참조)
pipeline {
    agent any

    options {
        disableConcurrentBuilds()   // 배포 중 또 배포 시작 금지
        timestamps()
    }

    environment {
        APP_DIR = '/srv/consol-app'
    }

    triggers {
        // 2분마다 git 변경 감지. webhook을 붙이면 이 블록은 제거해도 됨.
        pollSCM('H/2 * * * *')
    }

    stages {
        stage('Checkout') {
            steps { checkout scm }
        }

        stage('Syntax Check') {
            steps {
                // 전체 .py 컴파일 검사 — 문법 오류가 있으면 배포 자체를 중단
                sh 'docker run --rm -v "$WORKSPACE":/src -w /src python:3.12-slim python -m compileall -q .'
            }
        }

        stage('Deploy Code') {
            steps {
                // git이 추적하는 파일만 운영 디렉터리로 복사.
                // 서버 데이터(uploads/, *.json 상태 파일 등)는 절대 건드리지 않는다.
                sh '''
                    git ls-files > .deploy_files
                    rsync -a --files-from=.deploy_files ./ "$APP_DIR"/
                '''
            }
        }

        stage('Build & Restart') {
            steps {
                // build: requirements.txt/Dockerfile이 안 바뀌면 캐시로 수 초 내 완료
                // restart: 코드는 바인드 마운트라 재시작해야 반영됨 (수 초 단절 발생)
                sh '''
                    cd "$APP_DIR"
                    docker compose build
                    docker compose up -d
                    docker compose restart consol-web
                '''
            }
        }

        stage('Health Check') {
            steps {
                sh '''
                    for i in $(seq 1 15); do
                        if curl -fsS -o /dev/null http://127.0.0.1:5000/; then
                            echo "Health check OK"
                            exit 0
                        fi
                        sleep 2
                    done
                    echo "Health check 실패 — 컨테이너 로그:"
                    cd "$APP_DIR" && docker compose logs --tail 50
                    exit 1
                '''
            }
        }
    }

    post {
        failure {
            echo '배포 실패 — 위 콘솔 로그를 확인하세요. 서비스는 마지막 정상 상태로 남아있을 수 있으니 http://127.0.0.1:5000 응답 여부를 확인할 것.'
        }
    }
}
