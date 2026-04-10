# Usamos una imagen ligera de Python
FROM mcr.microsoft.com/devcontainers/python:3.12

# Instalamos uv directamente
COPY --from=ghcr.io/astral-sh/uv:latest /uv /uvx /bin/

# Evitamos que Python escriba archivos .pyc innecesarios
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Establecemos el directorio de trabajo dentro del contenedor
WORKDIR /workspaces/autoparser
USER vscode