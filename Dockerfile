FROM python:3.9-slim

# Instalar Chromium e dependências do sistema
# O comando clean limpa o cache para deixar a imagem leve
RUN apt-get update && apt-get install -y \
    chromium \
    chromium-driver \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY main.py .

# Comando de execução (sem buffer para logs em tempo real)
CMD ["python", "-u", "main.py"]
