// 연결 재무보고 통합 시스템 — 배포 파이프라인 (Jenkins 서버 → 운영서버 원격 배포)
// 전제 (배포 가이드 6장 참조):
//   - Jenkins는 젠킨스서버(176)에서 실행
//   - 젠킨스서버의 jenkins 계정이 운영서버(193)의 배포 계정으로 SSH 키 접속 가능
//   - 운영서버 배포 계정: docker 그룹 소속 + /opt/consol-app 소유
pipeline {
    agent any

    options {
        disableConcurrentBuilds()   // 배포 중 또 배포 시작 금지
        timestamps()
    }

    environment {
        // ★ 운영서버(193)의 실제 배포계정@IP로 수정할 것
        DEPLOY = 'saea@192.168.0.193'
        APP_DIR = '/opt/consol-app'
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
                // (젠킨스서버의 python3 사용. 없으면: sudo apt -y install python3)
                sh 'python3 -m compileall -q .'
            }
        }

        stage('Deploy Code') {
            steps {
                // git이 추적하는 파일만 운영서버로 전송.
                // 서버 데이터(uploads/, *.json 상태 파일 등)는 절대 건드리지 않는다.
                sh '''
                    git ls-files > .deploy_files
                    rsync -a --files-from=.deploy_files -e ssh ./ "$DEPLOY:$APP_DIR/"
                '''
            }
        }

        stage('Build & Restart') {
            steps {
                // build: requirements.txt/Dockerfile이 안 바뀌면 캐시로 수 초 내 완료
                // restart: 코드는 바인드 마운트라 재시작해야 반영됨 (수 초 단절 발생)
                sh '''
                    ssh "$DEPLOY" "cd $APP_DIR && docker compose build && docker compose up -d && docker compose restart consol-web"
                '''
            }
        }

        stage('Health Check') {
            steps {
                sh '''
                    for i in $(seq 1 15); do
                        if ssh "$DEPLOY" "curl -fsS -o /dev/null http://127.0.0.1:5000/"; then
                            echo "Health check OK"
                            exit 0
                        fi
                        sleep 2
                    done
                    echo "Health check 실패 — 운영서버 컨테이너 로그:"
                    ssh "$DEPLOY" "cd $APP_DIR && docker compose logs --tail 50"
                    exit 1
                '''
            }
        }
    }

    post {
        failure {
            echo '배포 실패 — 위 콘솔 로그를 확인하세요. 운영서버에서 docker compose ps 로 서비스 생존 여부를 확인할 것.'
        }
    }
}
