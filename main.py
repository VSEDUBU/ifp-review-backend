import os, httpx
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_origin_regex=".*",
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

GROQ_KEY = os.environ.get("GROQ_API_KEY", "")

# 依序嘗試：token 限制大的優先
MODELS = [
    "meta-llama/llama-4-scout-17b-16e-instruct",  # 30000 TPM
    "llama-3.3-70b-versatile",                     # 12000 TPM（備用）
]

class ReviewRequest(BaseModel):
    system: str
    user: str

@app.get("/")
def health():
    return {"status": "ok", "models": MODELS}

@app.post("/review")
async def review(req: ReviewRequest):
    if not GROQ_KEY:
        raise HTTPException(500, "Server: GROQ_API_KEY not set")

    last_error = ""
    async with httpx.AsyncClient(timeout=120) as client:
        for model in MODELS:
            resp = await client.post(
                "https://api.groq.com/openai/v1/chat/completions",
                headers={
                    "Authorization": f"Bearer {GROQ_KEY}",
                    "Content-Type": "application/json",
                },
                json={
                    "model": model,
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
                last_error = f"{model}: rate limited"
                continue  # 換下一個模型

            if resp.status_code != 200:
                last_error = f"{model}: {data.get('error', {}).get('message', resp.text[:200])}"
                continue

            text = data.get("choices", [{}])[0].get("message", {}).get("content", "")
            if text.strip():
                return {"result": text, "model": model}

            last_error = f"{model}: empty response"

    raise HTTPException(502, f"所有模型均失敗：{last_error}")
