# Imagen de Jenkins para el proyecto. Parte de la receta oficial de jenkins.io ("Installing Jenkins: Docker")
# y le agrega dos cosas que el pipeline necesita: el plugin de Docker Compose y permiso para leer el
# repositorio local que se monta en /repos.
FROM jenkins/jenkins:2.580.2-jdk21
USER root
RUN apt-get update && apt-get install -y lsb-release
RUN curl -fsSLo /usr/share/keyrings/docker-archive-keyring.asc \
  https://download.docker.com/linux/debian/gpg
RUN echo "deb [arch=$(dpkg --print-architecture) \
  signed-by=/usr/share/keyrings/docker-archive-keyring.asc] \
  https://download.docker.com/linux/debian \
  $(lsb_release -cs) stable" > /etc/apt/sources.list.d/docker.list
# docker-ce-cli es el cliente de Docker (receta oficial); docker-compose-plugin agrega "docker compose".
RUN apt-get update && apt-get install -y docker-ce-cli docker-compose-plugin
USER jenkins
# El repositorio montado pertenece a otro usuario del equipo anfitrión; sin esto git lo rechaza.
RUN git config --global --add safe.directory '*'
RUN jenkins-plugin-cli --plugins "docker-workflow json-path-api"
