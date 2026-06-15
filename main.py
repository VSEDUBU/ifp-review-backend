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

CEREBRAS_KEY = os.environ.get("CEREBRAS_API_KEY", "")
GROQ_KEY     = os.environ.get("GROQ_API_KEY", "")

# Cerebras 主力（1M tokens/day，60K TPM），Groq 備用（100K/day，12K TPM）
PROVIDERS = [
    {
        "name": "cerebras/llama-4-scout",
        "url": "https://api.cerebras.ai/v1/chat/completions",
        "model": "llama-4-scout",
        "key_env": "CEREBRAS",
    },
    {
        "name": "cerebras/qwen3-32b",
        "url": "https://api.cerebras.ai/v1/chat/completions",
        "model": "qwen-3-32b",
        "key_env": "CEREBRAS",
    },
    {
        "name": "groq/llama-3.3-70b",
        "url": "https://api.groq.com/openai/v1/chat/completions",
        "model": "llama-3.3-70b-versatile",
        "key_env": "GROQ",
    },
]

class ReviewRequest(BaseModel):
    system: str
    user: str

@app.get("/")
def health():
    return {"status": "ok", "providers": [p["name"] for p in PROVIDERS]}

@app.post("/review")
async def review(req: ReviewRequest):
    last_error = ""

    async with httpx.AsyncClient(timeout=120) as client:
        for p in PROVIDERS:
            key = CEREBRAS_KEY if p["key_env"] == "CEREBRAS" else GROQ_KEY
            if not key:
                last_error = f"{p['name']}: API key not set"
                continue

            try:
                resp = await client.post(
                    p["url"],
                    headers={
                        "Authorization": f"Bearer {key}",
                        "Content-Type": "application/json",
                    },
                    json={
                        "model": p["model"],
                        "max_tokens": 4096,
                        "temperature": 0.1,
                        "messages": [
                            {"role": "system", "content": req.system},
                            {"role": "user",   "content": req.user},
                        ],
                    },
                )
                data = resp.json()

                if resp.status_code == 429:
                    last_error = f"{p['name']}: rate limited"
                    continue

                if resp.status_code != 200:
                    last_error = f"{p['name']}: {data.get('error', {}).get('message', resp.text[:200])}"
                    continue

                text = data.get("choices", [{}])[0].get("message", {}).get("content", "")
                if text.strip():
                    return {"result": text, "model": p["name"]}

                last_error = f"{p['name']}: empty response"

            except Exception as e:
                last_error = f"{p['name']}: {str(e)}"

    raise HTTPException(502, f"所有模型均失敗：{last_error}")
