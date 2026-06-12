// 연결 재무보고 통합 시스템 — 배포 파이프라인 (원격 배포 버전)
// 전제: 176번 Jenkins 서버에서 193번 운영 서버로 SSH 접속이 가능해야 함.
//       Jenkins Credentials에 'consol-prod-ssh' ID로 SSH Private Key가 등록되어 있어야 함.
//       운영 서버의 대상 디렉터리(/srv/consol-app)에 TARGET_USER의 쓰기 권한이 있어야 함.
pipeline {
    agent any

    options {
        disableConcurrentBuilds()   // 배포 중 또 배포 시작 금지
        timestamps()
    }

    environment {
        TARGET_HOST = '192.168.0.193'         // 운영 서버 IP (운영 환경에 맞게 변경 가능)
        TARGET_USER = 'saea'                // 운영 서버 접속 계정
        APP_DIR = '/opt/consol-app'           // 운영 서버 내 애플리케이션 경로
    }

    stages {
        stage('Checkout') {
            steps { checkout scm }
        }

        stage('Syntax Check') {
            steps {
                // 전체 .py 컴파일 검사 — 문법 오류가 있으면 배포 자체를 중단 (Jenkins 서버의 python3 사용)
                sh 'python3 -m compileall -q .'
            }
        }

        stage('Deploy Code') {
            steps {
                // git이 추적하는 파일만 운영 서버의 디렉터리로 복사.
                // rsync over SSH 사용.
                sshagent(credentials: ['operating-server-ssh']) {
                    sh '''
                        git ls-files > .deploy_files
                        rsync -avz -e "ssh -o StrictHostKeyChecking=no" --files-from=.deploy_files ./ ${TARGET_USER}@${TARGET_HOST}:${APP_DIR}/
                    '''
                }
            }
        }

        stage('Build & Restart') {
            steps {
                // 원격 서버에서 docker compose 빌드 및 재시작 실행
                sshagent(credentials: ['operating-server-ssh']) {
                    sh '''
                        ssh -o StrictHostKeyChecking=no ${TARGET_USER}@${TARGET_HOST} "
                            cd ${APP_DIR}
                            docker compose build
                            docker compose up -d
                            docker compose restart consol-web
                        "
                    '''
                }
            }
        }

        stage('Health Check') {
            steps {
                // 원격 서버 자체에서 헬스체크를 수행하도록 명령 전송
                sshagent(credentials: ['operating-server-ssh']) {
                    sh '''
                        ssh -o StrictHostKeyChecking=no ${TARGET_USER}@${TARGET_HOST} '
                            for i in $(seq 1 15); do
                                if curl -fsS -o /dev/null http://127.0.0.1:5000/; then
                                    echo "Health check OK"
                                    exit 0
                                fi
                                sleep 2
                            done
                            echo "Health check 실패 — 컨테이너 로그:"
                            cd '"${APP_DIR}"' && docker compose logs --tail 50
                            exit 1
                        '
                    '''
                }
            }
        }
    }

    post {
        failure {
            echo '배포 실패 — 위 콘솔 로그를 확인하세요. 서비스는 마지막 정상 상태로 남아있을 수 있으니 운영서버 응답 여부를 확인할 것.'
        }
    }
}
