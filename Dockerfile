FROM python:3.11-slim

# Unbuffered output so print/logging reaches Render's log stream immediately.
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1

WORKDIR /app

# Install dependencies before copying the app, so editing a template or
# main.py reuses the cached install layer instead of reinstalling everything.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY main.py .
COPY templates/ templates/

# Render injects PORT; the fallback keeps `docker run -p 10000:10000` working
# locally. One worker fits the free tier's 512 MB, and the default 30s timeout
# is too tight for a long PDF. Shell form is required to expand ${PORT}.
CMD gunicorn main:app --bind 0.0.0.0:${PORT:-10000} --workers 1 --timeout 120
