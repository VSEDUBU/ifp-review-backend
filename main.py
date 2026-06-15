import os, httpx
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
)

GROQ_KEY   = os.environ.get("GROQ_API_KEY", "")
GROQ_MODEL = "llama-3.3-70b-versatile"

class ReviewRequest(BaseModel):
    system: str
    user: str

@app.get("/")
def health():
    return {"status": "ok", "model": GROQ_MODEL}

@app.post("/review")
async def review(req: ReviewRequest):
    if not GROQ_KEY:
        raise HTTPException(500, "Server: GROQ_API_KEY not set")

    async with httpx.AsyncClient(timeout=120) as client:
        resp = await client.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {GROQ_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": GROQ_MODEL,
                "max_tokens": 8192,
                "temperature": 0.1,
                "messages": [
                    {"role": "system", "content": req.system},
                    {"role": "user",   "content": req.user},
                ],
            },
        )

    data = resp.json()

    if resp.status_code == 429:
        raise HTTPException(429, "Groq 每分鐘額度已用完，請稍後再試。")
    if not resp.ok:
        err = data.get("error", {}).get("message", resp.text[:200])
        raise HTTPException(502, f"Groq API 錯誤：{err}")

    text = data.get("choices", [{}])[0].get("message", {}).get("content", "")
    if not text.strip():
        raise HTTPException(502, "Groq 回傳空白結果，請重試。")

    return {"result": text, "model": GROQ_MODEL}
