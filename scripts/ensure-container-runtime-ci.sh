#!/bin/sh

set -eu

if ! command -v docker >/dev/null 2>&1 && ! command -v podman >/dev/null 2>&1; then
  apt-get update
  apt-get install -y --no-install-recommends docker.io docker-compose-plugin || \
    apt-get install -y --no-install-recommends docker.io docker-compose
fi

if command -v docker >/dev/null 2>&1; then
  docker --version
  if docker compose version >/dev/null 2>&1; then
    docker compose version
  elif command -v docker-compose >/dev/null 2>&1; then
    docker-compose version
  fi
  if docker ps >/dev/null 2>&1; then
    exit 0
  fi
fi

if command -v podman >/dev/null 2>&1; then
  podman --version
  if podman compose version >/dev/null 2>&1; then
    podman compose version
  elif command -v podman-compose >/dev/null 2>&1; then
    podman-compose version
  fi
  if podman ps >/dev/null 2>&1; then
    exit 0
  fi
fi

echo "No usable container runtime is available for the browser stack."
exit 1
