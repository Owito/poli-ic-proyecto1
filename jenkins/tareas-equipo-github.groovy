// Script del trabajo "tareas-equipo-github" (definido en Jenkins como "Pipeline script").
// Comprueba la conexión entre Jenkins y GitHub: descarga la rama main del repositorio público y corre las
// pruebas unitarias. La rama main todavía no tiene Jenkinsfile; por eso el script vive en el trabajo.
pipeline {
    agent any
    options {
        timestamps()
        timeout(time: 10, unit: 'MINUTES')
    }
    stages {
        stage('Descargar desde GitHub') {
            steps {
                git url: 'https://github.com/Owito/poli-ic-proyecto1.git', branch: 'main'
                sh 'git log -1 --format="Commit %h de %an: %s"'
            }
        }
        stage('Pruebas unitarias') {
            agent {
                docker {
                    image 'node:22-alpine'
                    reuseNode true
                }
            }
            steps {
                sh 'npm ci --cache .npm-cache --no-audit --no-fund'
                timeout(time: 3, unit: 'MINUTES') {
                    sh 'npm test'
                }
            }
        }
    }
}
