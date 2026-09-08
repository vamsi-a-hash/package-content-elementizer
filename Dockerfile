FROM python:3.13-slim-bookworm

RUN apt-get update && apt-get install -y curl 
RUN apt-get install -y --no-install-recommends libreoffice-writer \
    fonts-crosextra-carlito fonts-crosextra-caladea fonts-liberation2 \
    && rm -rf /var/lib/apt/lists/*
RUN curl -sSL https://install.python-poetry.org | python3 -

ENV PATH="/root/.local/bin:$PATH"

WORKDIR /app
COPY pyproject.toml poetry.lock ./

RUN --mount=type=cache,target=/root/.cache/pypoetry \
    --mount=type=cache,target=/root/.cache/pip \
    poetry config virtualenvs.create false && poetry lock && poetry install --no-root --no-interaction --no-ansi --only main

COPY ./talkingdb_ce /app/talkingdb_ce

ENV PYTHONUNBUFFERED=1 \
    PYTHONPATH=/app

COPY ./README.md /README.md
COPY ./version.txt /version.txt

EXPOSE 8091

CMD ["poetry", "run", "uvicorn", "talkingdb_ce.main:app", "--host", "0.0.0.0", "--port", "8091", "--workers", "1", "--loop", "uvloop", "--http", "httptools"]
