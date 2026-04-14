from __future__ import annotations

import json
import os
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path
import secrets
import string
from typing import Any, Dict

import boto3

PROJECT_ROOT = Path(__file__).resolve().parents[2]
ORIGINAL_SCRIPT = PROJECT_ROOT / "app" / "pipeline" / "originals" / "colab201.py"
ORIGINAL_SCRIPT_131 = PROJECT_ROOT / "app" / "pipeline" / "originals" / "colab1-3-1.py"

# 期待する固定ファイル名（WORK_DIR 配下）
SPEC_FILENAME = "エクセル転記仕様.xlsx"
TEMPLATE_FILENAME = "CF付財務分析表（経営指標あり）_ReadingData.xlsx"

# zipアップロード時にファイル名が "#U...." 形式で入るケースに対応（UI側のエンコード表現）
SPEC_FILENAME_ALT = "#U30a8#U30af#U30bb#U30eb#U8ee2#U8a18#U4ed5#U69d8.xlsx"
TEMPLATE_FILENAME_ALT = (
    "CF#U4ed8#U8ca1#U52d9#U5206#U6790#U8868#Uff08#U7d4c#U55b6#U6307#U6a19#U3042#U308a#Uff09_ReadingData.xlsx"
)


def _run(cmd: list[str], cwd: Path, env: Dict[str, str]) -> str:
    p = subprocess.run(
        cmd,
        cwd=str(cwd),
        env=env,
        capture_output=True,
        text=True,
    )
    if p.returncode != 0:
        raise RuntimeError(
            "Command failed:\n"
            f"cmd={cmd}\n"
            f"returncode={p.returncode}\n"
            f"stdout:\n{p.stdout}\n"
            f"stderr:\n{p.stderr}\n"
        )
    return p.stdout


def _ensure_work_assets(work_dir: Path) -> None:
    """
    colab201.py が参照する /tmp/work 相当のディレクトリに、
    仕様ExcelとテンプレExcelを配置する。

    既定では /app/app/pipeline/assets/ 以下を探す（Dockerに同梱する想定）。
    """
    assets_dir = PROJECT_ROOT / "app" / "pipeline" / "assets"
    spec_src = assets_dir / SPEC_FILENAME
    tpl_src = assets_dir / TEMPLATE_FILENAME

    if not spec_src.exists():
        spec_src = assets_dir / SPEC_FILENAME_ALT
    if not tpl_src.exists():
        tpl_src = assets_dir / TEMPLATE_FILENAME_ALT

    missing = []
    if not spec_src.exists():
        missing.append(str(spec_src))
    if not tpl_src.exists():
        missing.append(str(tpl_src))
    if missing:
        raise FileNotFoundError(
            "必要なExcelテンプレ/仕様ファイルが見つかりませんでした。\n"
            "Dockerイメージに同梱するか、起動時に配置してください。\n"
            f"探した場所: {assets_dir}\n"
            f"不足: {missing}"
        )

    (work_dir / SPEC_FILENAME).write_bytes(spec_src.read_bytes())
    (work_dir / TEMPLATE_FILENAME).write_bytes(tpl_src.read_bytes())


def run_colab201(api_payload: Dict[str, Any]) -> Dict[str, Any]:
    """
    API入力:
      {"data":[...], "ai_case_id": 123, "mode": "201"}
    を受け取り、colab201.py 実行後に colab1-3-1.py を必ず実行して
    「更新済みExcel」を完成させ、S3に保存して署名付きURLを返す。
    """
    ai_case_id = api_payload.get("ai_case_id")

    # 1) 専用の作業ディレクトリ（同時実行でも衝突しない）
    run_dir = Path(tempfile.mkdtemp(prefix="cashai03_201_", dir="/tmp"))
    work_dir = run_dir / "work"
    work_dir.mkdir(parents=True, exist_ok=True)

    # 2) 必要なExcelファイルを配置（WORK_DIR配下）
    _ensure_work_assets(work_dir)

    # 3) 入力データを output_updated.json として保存（colab201.pyが読む）
    data = api_payload.get("data", api_payload)
    (work_dir / "output_updated.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    # 3.5) colab1-3-1.py 用の financial.json を保存（WORK_DIR配下）
    fin = api_payload.get("financial_response", None)
    if isinstance(fin, list):
        fin_obj: Any = {"response": fin}
    elif isinstance(fin, dict) and isinstance(fin.get("response"), list):
        fin_obj = fin
    else:
        # NULL / 空 / 想定外は null を書く（colab1-3-1.py 側で空配列扱い等にできる）
        fin_obj = None

    (work_dir / "financial.json").write_text(
        json.dumps(fin_obj, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    # 4) 実行環境
    env = dict(os.environ)
    env["WORK_DIR"] = str(work_dir)

    # 5) 実行（1/2）: Excel 転記
    _run(["python3", str(ORIGINAL_SCRIPT)], cwd=work_dir, env=env)

    # 5) 実行（2/2）: 総合所見の自動生成（必ず実行）
    _run(["python3", str(ORIGINAL_SCRIPT_131)], cwd=work_dir, env=env)

    # 6) 出力を読み込み（Excel + log）
    out_excel = work_dir / "CF付財務分析表（経営指標あり）_ReadingData_updated.xlsx"
    out_log = work_dir / "transfer_log.txt"

    if not out_excel.exists():
        raise RuntimeError("更新済みExcelが生成されませんでした（colab201.py のログを確認してください）")

    log_text = out_log.read_text(encoding="utf-8", errors="replace") if out_log.exists() else ""

    # 7) S3へアップロードして署名付きURL返却
    bucket_name = os.environ.get("S3_BUCKET")
    region = os.environ.get("S3_REGION")
    access_key = os.environ.get("S3_ACCESS_KEY")
    secret_key = os.environ.get("S3_SECRET_KEY")
    expires = int(os.environ.get("PRESIGN_EXPIRES", "3600"))
    if not bucket_name or not region or not access_key or not secret_key:
        raise RuntimeError(
            "S3環境変数が未設定です。S3_BUCKET/S3_REGION/S3_ACCESS_KEY/S3_SECRET_KEY を設定してください。"
        )

    prefix = os.environ.get("S3_PREFIX", "cash-ai-03")

    ts = datetime.now().strftime("%Y%m%d%H%M%S")
    rand15 = "".join(secrets.choice(string.ascii_letters + string.digits) for _ in range(15))

    base_filename = f"CF付財務分析表_ai_case_{ai_case_id}_201" if ai_case_id else "CF付財務分析表_201"
    upload_filename = f"{base_filename}_{ts}_{rand15}.xlsx"
    object_name = f"{prefix}/{ai_case_id}/{upload_filename}" if ai_case_id else f"{prefix}/{upload_filename}"

    s3 = boto3.client(
        "s3",
        region_name=region,
        aws_access_key_id=access_key,
        aws_secret_access_key=secret_key,
    )

    s3.upload_file(
        str(out_excel),
        bucket_name,
        object_name,
        ExtraArgs={
            "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
    )

    signed_url = s3.generate_presigned_url(
        ClientMethod="get_object",
        Params={"Bucket": bucket_name, "Key": object_name},
        ExpiresIn=expires,
    )

    return {
        "runner": "runner201",
        "ai_case_id": ai_case_id,
        "excel_filename": upload_filename,
        "s3_bucket": bucket_name,
        "s3_region": region,
        "s3_key": object_name,
        "url": signed_url,
        "url_expires_in": expires,
        "transfer_log": log_text,
    }
