import os, httpx, logging
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

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

# gpt-oss-120b 是確認可用的，其他等查到正確 ID 再加
PROVIDERS = [
    {
        "name": "cerebras/gpt-oss-120b",
        "url": "https://api.cerebras.ai/v1/chat/completions",
        "model": "gpt-oss-120b",
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
    return {
        "status": "ok",
        "providers": [p["name"] for p in PROVIDERS],
        "cerebras_key_set": bool(CEREBRAS_KEY),
        "groq_key_set": bool(GROQ_KEY),
    }

@app.get("/models")
async def list_models():
    """查詢 Cerebras 實際可用的 model 清單"""
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.get(
            "https://api.cerebras.ai/v1/models",
            headers={"Authorization": f"Bearer {CEREBRAS_KEY}"},
        )
    return resp.json()

@app.post("/review")
async def review(req: ReviewRequest):
    last_error = ""

    async with httpx.AsyncClient(timeout=120) as client:
        for p in PROVIDERS:
            key = CEREBRAS_KEY if p["key_env"] == "CEREBRAS" else GROQ_KEY
            if not key:
                last_error = f"{p['name']}: API key not set"
                logger.warning(last_error)
                continue

            try:
                logger.info(f"Trying {p['name']}...")
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
                logger.info(f"{p['name']} status: {resp.status_code}")

                if resp.status_code == 429:
                    last_error = f"{p['name']}: rate limited"
                    logger.warning(last_error)
                    continue

                if resp.status_code != 200:
                    err_msg = data.get('error', {}).get('message', str(data)[:300])
                    last_error = f"{p['name']}: HTTP {resp.status_code} - {err_msg}"
                    logger.error(last_error)
                    continue

                text = data.get("choices", [{}])[0].get("message", {}).get("content", "")
                if text.strip():
                    logger.info(f"Success with {p['name']}")
                    return {"result": text, "model": p["name"]}

                last_error = f"{p['name']}: empty response"
                logger.warning(last_error)

            except Exception as e:
                last_error = f"{p['name']}: {str(e)}"
                logger.error(last_error)

    raise HTTPException(502, f"所有模型均失敗：{last_error}")
