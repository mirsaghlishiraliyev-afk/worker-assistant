#!/usr/bin/env sh
# TeamCity build addımı (Command Line runner: "sh deploy/teamcity_build.sh"), test.server-də işləyən agentdə.
# TeamCity BUILD_NUMBER mühit dəyişənini özü verir. Stack adı serverdəki stack ilə eyni olmalıdır (docker stack ls).
set -e
TEQ="${BUILD_NUMBER:-latest}"
STACK="${WA_STACK:-asis_dev}"
SERVIS="${STACK}_worker_assistant"

# Testlər Dockerfile-ın içində işləyir: test sınsa, build burada dayanır
docker build -t "worker-assistant-dev-image:${TEQ}" -t worker-assistant-dev-image:latest .

if docker service inspect "$SERVIS" >/dev/null 2>&1; then
  docker service update --force --image "worker-assistant-dev-image:${TEQ}" "$SERVIS"
else
  echo "«$SERVIS» servisi hələ yoxdur. Serverdəki stack faylına deploy/docker-compose.worker-assistant.yml-dakı servisi əlavə edib"
  echo "bir dəfə işə salın:  WA_PAROL=... docker stack deploy -c <stack faylı> $STACK"
  exit 1
fi
