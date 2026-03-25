FROM lordjabez/claude-code:latest

COPY pyproject.toml uv.lock ./

RUN uv sync

COPY hooks/ /home/claude/hooks/
