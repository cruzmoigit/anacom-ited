FROM python:3.11-slim

WORKDIR /app

# Dependências do sistema (mínimas)
RUN apt-get update && apt-get install -y --no-install-recommends \
    ca-certificates \
 && rm -rf /var/lib/apt/lists/*

# Dependências Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Código da aplicação
COPY server.py distritos.json ./
COPY static ./static

# Porta exposta
EXPOSE 8765

# Variável de ambiente para cache (opcional)
ENV PYTHONUNBUFFERED=1

CMD ["python", "server.py"]
