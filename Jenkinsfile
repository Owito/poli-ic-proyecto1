// Pipeline de integración continua de la API de tareas (Entrega 2).
// Cada cambio en la rama se descarga, se prueba, se construye en contenedores y se verifica de punta a punta.
pipeline {
    agent any

    options {
        timestamps()
        disableConcurrentBuilds()
        // Si algo se queda esperando (por ejemplo, una prueba que no termina), la ejecución se corta sola.
        timeout(time: 15, unit: 'MINUTES')
        buildDiscarder(logRotator(numToKeepStr: '30'))
    }

    // Jenkins revisa el repositorio cada dos minutos y lanza una ejecución si encuentra commits nuevos.
    triggers {
        pollSCM('H/2 * * * *')
    }

    environment {
        COMPOSE_PROJECT_NAME = 'tareas-ci'
        APP_PUERTO = '8095'
    }

    stages {
        stage('Descargar código') {
            steps {
                checkout scm
                sh 'git log -1 --format="Commit %h de %an: %s"'
            }
        }

        stage('Instalar dependencias y pruebas unitarias') {
            // Node.js no está instalado en Jenkins: la etapa corre dentro de un contenedor node:22-alpine.
            agent {
                docker {
                    image 'node:22-alpine'
                    reuseNode true
                }
            }
            steps {
                sh 'node --version && npm --version'
                sh 'npm ci --cache .npm-cache --no-audit --no-fund'
                timeout(time: 3, unit: 'MINUTES') {
                    sh 'npm test'
                }
            }
        }

        stage('Construir la imagen de la API') {
            steps {
                sh 'docker compose build app'
            }
        }

        stage('Levantar los dos contenedores') {
            steps {
                // La clave de la base de datos sale de una credencial de Jenkins; nunca queda en el repositorio.
                withCredentials([string(credentialsId: 'tareas-db-password', variable: 'CLAVE_DB')]) {
                    sh '''
                        cat > .env <<EOF
POSTGRES_USER=tareas_user
POSTGRES_PASSWORD=${CLAVE_DB}
POSTGRES_DB=tareas
APP_PUERTO=${APP_PUERTO}
EOF
                        docker compose up -d --wait
                        docker compose ps
                    '''
                }
            }
        }

        stage('Verificar /salud') {
            steps {
                // "docker" es el nombre del motor de Docker en la red de Jenkins; ahí quedó publicado el puerto 8095.
                sh 'curl -fsS http://docker:${APP_PUERTO}/salud'
            }
        }

        stage('Pruebas de integración') {
            steps {
                sh 'docker compose exec -T app npm test'
            }
        }
    }

    post {
        always {
            sh 'docker compose down -v --remove-orphans || true'
            sh 'rm -f .env'
        }
        success {
            echo 'Integración verificada: pruebas, imagen, contenedores y /salud en verde.'
        }
        failure {
            echo 'La integración falló: revise la etapa en rojo antes de seguir integrando cambios.'
        }
    }
}
