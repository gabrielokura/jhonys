from fastapi import FastAPI, HTTPException

from .db import ping_firebird

app = FastAPI(title="Relatorios API")


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/db/ping")
def db_ping() -> dict[str, str | int]:
    try:
        result = ping_firebird()
    except Exception as exc:
        raise HTTPException(
            status_code=503,
            detail=f"Firebird ping failed: {exc}",
        ) from exc

    return {"status": "ok", "result": result}
