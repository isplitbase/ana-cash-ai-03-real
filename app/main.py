from fastapi import FastAPI, Body
from typing import Any, Dict

from app.pipeline.runner201 import run_colab201
from app.pipeline.runner202 import run_colab202

app = FastAPI()

from fastapi.middleware.cors import CORSMiddleware

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],          # とりあえず検証用。本番はドメインを絞るの推奨
    allow_credentials=False,      # "*" のときは False 推奨
    allow_methods=["*"],          # OPTIONS/POST などを許可
    allow_headers=["*"],          # Content-Type など
)

@app.get("/health")
def health():
    return {"ok": True}


@app.post("/v1/pipeline")
def pipeline(payload: Dict[str, Any] = Body(...)):
    """
    入力例:
      {
        "data": [...],
        "ai_case_id": 21160,
        "mode": "201" | "202" | "both"   # 任意（省略時 201）
      }

    返却:
      mode=201/202 の場合: {"result": {"url": "...", ...}}
      mode=both の場合: {"results": [ {"url": "..."}, {"url": "..."} ]}
    """
    # 既定は colab201 のみ（運用向け）
    mode = str(payload.get("mode", "201")).lower()

    if mode in ("201", "colab201"):
        return {"result": run_colab201(payload)}
    if mode in ("202", "colab202"):
        return {"result": run_colab202(payload)}

    # both: 両方実行（差分比較にも使える）
    return {"results": [run_colab201(payload), run_colab202(payload)]}
