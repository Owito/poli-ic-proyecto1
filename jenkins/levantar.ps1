# Levanta Jenkins en Docker siguiendo la guía oficial de jenkins.io para Windows, con tres ajustes:
#   1) el puerto web se publica en el 8181 del equipo, porque el 8080 ya lo usa otra aplicación;
#   2) se monta en /repos una copia local (bare) del repositorio, para que Jenkins lea la rama de trabajo
#      antes de que se publique en GitHub;
#   3) se habilita la lectura de ese repositorio local con la propiedad ALLOW_LOCAL_CHECKOUT del plugin Git.
# Uso: powershell -File jenkins\levantar.ps1 -Espejo C:\ruta\a\repo.git
param(
    [Parameter(Mandatory = $true)][string]$Espejo,
    [int]$PuertoWeb = 8181
)
$ErrorActionPreference = 'Stop'
$aqui = $PSScriptRoot

# 1. Red propia para Jenkins y el motor de Docker que usará.
if (-not (docker network ls --format '{{.Name}}' | Select-String -SimpleMatch -Pattern 'jenkins' -Quiet)) {
    docker network create jenkins
}

# 2. Motor de Docker dentro de Docker (docker:dind). Comparte el volumen jenkins-data para que los
#    contenedores que lance el pipeline vean los archivos del espacio de trabajo.
docker run --name jenkins-docker --rm --detach `
  --privileged --network jenkins --network-alias docker `
  --env DOCKER_TLS_CERTDIR=/certs `
  --volume jenkins-docker-certs:/certs/client `
  --volume jenkins-data:/var/jenkins_home `
  --publish 2376:2376 `
  docker:dind

# 3. Imagen de Jenkins con el cliente de Docker.
docker build -t tareas-jenkins:2.580.2-1 $aqui

# 4. Contenedor de Jenkins.
docker run --name jenkins --restart=on-failure --detach `
  --network jenkins --env DOCKER_HOST=tcp://docker:2376 `
  --env DOCKER_CERT_PATH=/certs/client --env DOCKER_TLS_VERIFY=1 `
  --env JAVA_OPTS="-Dhudson.plugins.git.GitSCM.ALLOW_LOCAL_CHECKOUT=true" `
  --volume jenkins-data:/var/jenkins_home `
  --volume jenkins-docker-certs:/certs/client:ro `
  --volume "${Espejo}:/repos/poli-ic-proyecto1.git:ro" `
  --publish "${PuertoWeb}:8080" --publish 50000:50000 `
  tareas-jenkins:2.580.2-1

Write-Output "Jenkins en http://localhost:$PuertoWeb (la clave inicial está en /var/jenkins_home/secrets/initialAdminPassword)"
