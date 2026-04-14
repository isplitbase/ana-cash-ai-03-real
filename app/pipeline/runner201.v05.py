from __future__ import annotations

import json
import os
import secrets
import shutil
import string
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Tuple

import boto3

PROJECT_ROOT = Path(__file__).resolve().parents[2]

# Originals
ORIGINAL_SCRIPT_201 = PROJECT_ROOT / "app" / "pipeline" / "originals" / "colab201.py"
ORIGINAL_SCRIPT_131 = PROJECT_ROOT / "app" / "pipeline" / "originals" / "colab1-3-1.py"
ORIGINAL_SCRIPT_15 = PROJECT_ROOT / "app" / "pipeline" / "originals" / "colab1-5.py"

# 期待する固定ファイル名（WORK_DIR 配下）
SPEC_FILENAME = "エクセル転記仕様.xlsx"
TEMPLATE_FILENAME = "CF付財務分析表（経営指標あり）_ReadingData.xlsx"
CF_TEMPLATE_FILENAME = "CF資金移動表.xlsx"

# zipアップロード時にファイル名が "#U...." 形式で入るケースに対応（UI側のエンコード表現）
SPEC_FILENAME_ALT = "#U30a8#U30af#U30bb#U30eb#U8ee2#U8a18#U4ed5#U69d8.xlsx"
TEMPLATE_FILENAME_ALT = (
    "CF#U4ed8#U8ca1#U52d9#U5206#U6790#U8868#Uff08#U7d4c#U55b6#U6307#U6a19#U3042#U308a#Uff09_ReadingData.xlsx"
)
CF_TEMPLATE_FILENAME_ALT = "CF#U8cc7#U91d1#U79fb#U52d5#U8868.xlsx"


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
    colab201.py / colab1-5.py が参照する作業ディレクトリに、
    仕様Excel・テンプレExcel類を配置する。

    既定では /app/app/pipeline/assets/ 以下を探す（Dockerに同梱する想定）。
    """
    assets_dir = PROJECT_ROOT / "app" / "pipeline" / "assets"

    spec_src = assets_dir / SPEC_FILENAME
    if not spec_src.exists():
        spec_src = assets_dir / SPEC_FILENAME_ALT

    tpl_src = assets_dir / TEMPLATE_FILENAME
    if not tpl_src.exists():
        tpl_src = assets_dir / TEMPLATE_FILENAME_ALT

    cf_src = assets_dir / CF_TEMPLATE_FILENAME
    if not cf_src.exists():
        cf_src = assets_dir / CF_TEMPLATE_FILENAME_ALT

    missing = []
    if not spec_src.exists():
        missing.append(str(spec_src))
    if not tpl_src.exists():
        missing.append(str(tpl_src))
    if not cf_src.exists():
        missing.append(str(cf_src))

    if missing:
        raise FileNotFoundError(
            "必要なExcelテンプレ/仕様ファイルが見つかりませんでした。\n"
            "Dockerイメージに同梱するか、起動時に配置してください。\n"
            f"探した場所: {assets_dir}\n"
            f"不足: {missing}"
        )

    (work_dir / SPEC_FILENAME).write_bytes(spec_src.read_bytes())
    (work_dir / TEMPLATE_FILENAME).write_bytes(tpl_src.read_bytes())
    (work_dir / CF_TEMPLATE_FILENAME).write_bytes(cf_src.read_bytes())


def _s3_client() -> Tuple[Any, str, str, str, int]:
    bucket_name = os.environ.get("S3_BUCKET")
    region = os.environ.get("S3_REGION")
    access_key = os.environ.get("S3_ACCESS_KEY")
    secret_key = os.environ.get("S3_SECRET_KEY")
    # 既定は 7時間 (25200秒)
    expires = int(os.environ.get("PRESIGN_EXPIRES", "25200"))
    prefix = os.environ.get("S3_PREFIX", "cash-ai-03")

    if not bucket_name or not region or not access_key or not secret_key:
        raise RuntimeError(
            "S3環境変数が未設定です。S3_BUCKET/S3_REGION/S3_ACCESS_KEY/S3_SECRET_KEY を設定してください。"
        )

    s3 = boto3.client(
        "s3",
        region_name=region,
        aws_access_key_id=access_key,
        aws_secret_access_key=secret_key,
    )
    return s3, bucket_name, region, prefix, expires


def _upload_and_presign(
    s3: Any,
    bucket_name: str,
    region: str,
    prefix: str,
    expires: int,
    local_path: Path,
    ai_case_id: Any,
    base_filename: str,
) -> Dict[str, Any]:
    ts = datetime.now().strftime("%Y%m%d%H%M%S")
    rand15 = "".join(secrets.choice(string.ascii_letters + string.digits) for _ in range(15))
    upload_filename = f"{base_filename}_{ts}_{rand15}.xlsx"
    object_name = f"{prefix}/{ai_case_id}/{upload_filename}" if ai_case_id else f"{prefix}/{upload_filename}"

    s3.upload_file(
        str(local_path),
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
        "excel_filename": upload_filename,
        "s3_bucket": bucket_name,
        "s3_region": region,
        "s3_key": object_name,
        "url": signed_url,
        "url_expires_in": expires,
    }


def run_colab201(api_payload: Dict[str, Any]) -> Dict[str, Any]:
    """
    API入力:
      {"data":[...], "ai_case_id": 123, "mode": "201"}
    を受け取り、

    1) colab201.py 実行（ReadingData_updated 生成）
    2) colab1-3-1.py 実行（総合所見）
    3) colab1-5.py 実行（CF資金移動表_updated 生成）
    4) それぞれS3にアップロードし、署名付きURL（既定7時間）を返す

    ※重要：既存の返却キー（トップレベル）を変更しない
      - excel_filename / s3_bucket / s3_region / s3_key / url / url_expires_in / transfer_log は従来通り
      - 追加で excel1_5 を返す
    """
    ai_case_id = api_payload.get("ai_case_id")

    # 同時実行でも衝突しない作業ディレクトリ
    run_dir = Path(tempfile.mkdtemp(prefix="cashai03_201_", dir="/tmp"))
    work_dir = run_dir / "work"

    try:
        work_dir.mkdir(parents=True, exist_ok=True)

        # 必要なExcelファイルを配置（WORK_DIR配下）
        _ensure_work_assets(work_dir)

        # 入力データを output_updated.json として保存（colab201.pyが読む）
        data = api_payload.get("data", api_payload)
        (work_dir / "output_updated.json").write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        # colab1-3-1.py 用の financial.json を保存（WORK_DIR配下）
        fin = api_payload.get("financial_response", None)
        if isinstance(fin, list):
            fin_obj: Any = {"response": fin}
        elif isinstance(fin, dict) and isinstance(fin.get("response"), list):
            fin_obj = fin
        else:
            fin_obj = None

        (work_dir / "financial.json").write_text(
            json.dumps(fin_obj, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        # 実行環境
        env = dict(os.environ)
        env["WORK_DIR"] = str(work_dir)

        # 1) colab201.py
        _run(["python3", str(ORIGINAL_SCRIPT_201)], cwd=work_dir, env=env)

        # 2) colab1-3-1.py（kousya の場合のみ実行）
        kousya_flag = None
        if isinstance(data, dict):
            kousya_flag = data.get("kousya")

        if kousya_flag == "kousya" and ORIGINAL_SCRIPT_131.exists():
            _run(["python3", str(ORIGINAL_SCRIPT_131)], cwd=work_dir, env=env)


        # 3) colab1-5.py
        _run(["python3", str(ORIGINAL_SCRIPT_15)], cwd=work_dir, env=env)

        # 出力ファイル
        out_excel_201 = work_dir / "CF付財務分析表（経営指標あり）_ReadingData_updated.xlsx"
        out_cf = work_dir / "CF資金移動表_updated.xlsx"
        out_log = work_dir / "transfer_log.txt"

        if not out_excel_201.exists():
            raise RuntimeError("更新済みExcelが生成されませんでした（colab201.py のログを確認してください）")
        if not out_cf.exists():
            raise RuntimeError("CF資金移動表の更新Excelが生成されませんでした（colab1-5.py のログを確認してください）")

        log_text = out_log.read_text(encoding="utf-8", errors="replace") if out_log.exists() else ""

        # S3へアップロードして署名付きURL返却（既定7時間）
        s3, bucket_name, region, prefix, expires = _s3_client()

        base1 = f"CF付財務分析表_ai_case_{ai_case_id}_201" if ai_case_id else "CF付財務分析表_201"
        base2 = f"CF資金移動表_ai_case_{ai_case_id}_201" if ai_case_id else "CF資金移動表_201"

        up1 = _upload_and_presign(
            s3=s3,
            bucket_name=bucket_name,
            region=region,
            prefix=prefix,
            expires=expires,
            local_path=out_excel_201,
            ai_case_id=ai_case_id,
            base_filename=base1,
        )
        up2 = _upload_and_presign(
            s3=s3,
            bucket_name=bucket_name,
            region=region,
            prefix=prefix,
            expires=expires,
            local_path=out_cf,
            ai_case_id=ai_case_id,
            base_filename=base2,
        )

        # ★既存返却を維持（トップレベルは up1 で埋める）
        return {
            "runner": "runner201",
            "ai_case_id": ai_case_id,
            "excel_filename": up1["excel_filename"],
            "s3_bucket": up1["s3_bucket"],
            "s3_region": up1["s3_region"],
            "s3_key": up1["s3_key"],
            "url": up1["url"],
            "url_expires_in": up1["url_expires_in"],
            "transfer_log": log_text,

            # ★追加：colab1-5.py 側（既存キーは変えず、追加するだけ）
            "excel1_5": {
                "excel_filename": up2["excel_filename"],
                "s3_bucket": up2["s3_bucket"],
                "s3_region": up2["s3_region"],
                "s3_key": up2["s3_key"],
                "url": up2["url"],
                "url_expires_in": up2["url_expires_in"],
            },
        }

    finally:
        shutil.rmtree(run_dir, ignore_errors=True)
