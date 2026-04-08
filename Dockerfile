FROM python:3.12-slim

WORKDIR /app

# System deps for pdfplumber (Pillow needs some libs)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libglib2.0-0 libsm6 libxrender1 libxext6 \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY src/ src/

# Create output/downloads dirs
RUN mkdir -p output downloads

# Expose dashboard port
EXPOSE 8000

# Run dashboard with uvicorn
CMD ["uvicorn", "src.dashboard:app", "--host", "0.0.0.0", "--port", "8000"]
