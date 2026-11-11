# Converteme — servidor de conversão (Google Cloud Run)
FROM python:3.12-slim

# LibreOffice só com o necessário (Word, Excel, PowerPoint e o Draw, que abre PDF para o
# PDF -> PowerPoint), Ghostscript (PDF/A) e fontes livres com as
# mesmas medidas das fontes do Office (Carlito = Calibri, Caladea = Cambria,
# Liberation = Arial / Times New Roman / Courier New)
RUN apt-get update \
 && apt-get install -y --no-install-recommends \
      libreoffice-writer-nogui libreoffice-calc-nogui libreoffice-impress-nogui libreoffice-draw-nogui ghostscript \
      fonts-crosextra-carlito fonts-crosextra-caladea fonts-liberation2 fonts-dejavu-core fonts-noto-core \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app.py layout.py planilha.py ./

# usuário sem privilégios
RUN useradd --create-home --uid 10001 conversor
USER conversor

ENV PORT=8080 PYTHONUNBUFFERED=1
EXPOSE 8080
CMD ["python", "app.py"]
