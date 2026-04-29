from __future__ import annotations

from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from jarvis_core import JarvisAssistant

BASE_DIR = Path(__file__).resolve().parent
WEB_DIR = BASE_DIR / "web"

app = FastAPI(title="Jarvis Web Assistant", version="1.0.0")
assistant = JarvisAssistant(model="phi", enable_tts=False)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/static", StaticFiles(directory=str(WEB_DIR)), name="static")


class ChatRequest(BaseModel):
    message: str


@app.get("/api/health")
def health() -> dict:
    return {"status": "ok", "model": assistant.model}


@app.post("/api/chat")
def chat(payload: ChatRequest) -> dict:
    result = assistant.process_text(payload.message)
    return {"reply": result.text, "action": result.action}


@app.get("/")
def index() -> FileResponse:
    return FileResponse(WEB_DIR / "index.html")

