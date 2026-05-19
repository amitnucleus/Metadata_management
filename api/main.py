import io
import json

import pandas as pd
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(title="Metadata Parser API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


def df_to_sheet(df: pd.DataFrame) -> dict:
    df = df.where(pd.notnull(df), None)
    headers = list(df.columns.astype(str))
    rows = [
        {h: (None if v is None else v) for h, v in zip(headers, row)}
        for row in df.itertuples(index=False, name=None)
    ]
    return {"headers": headers, "rows": rows}


@app.post("/parse")
async def parse_file(file: UploadFile = File(...)):
    name = file.filename or ""
    ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
    raw = await file.read()

    try:
        if ext in ("xlsx", "xls", "ods"):
            wb = pd.read_excel(io.BytesIO(raw), sheet_name=None, dtype=str)
            return {sheet: df_to_sheet(df) for sheet, df in wb.items()}

        if ext == "csv":
            text = raw.decode("utf-8-sig")
            df = pd.read_csv(io.StringIO(text), dtype=str)
            sheet_name = name.rsplit(".", 1)[0] or "Sheet1"
            return {sheet_name: df_to_sheet(df)}

        if ext == "json":
            data = json.loads(raw)
            if isinstance(data, list):
                df = pd.json_normalize(data).astype(str)
                sheet_name = name.rsplit(".", 1)[0] or "Sheet1"
                return {sheet_name: df_to_sheet(df)}
            if isinstance(data, dict):
                result = {}
                for sheet, value in data.items():
                    if isinstance(value, list):
                        df = pd.json_normalize(value).astype(str)
                        result[sheet] = df_to_sheet(df)
                    else:
                        result[sheet] = {"headers": [], "rows": []}
                return result
            raise ValueError("JSON must be an array or object")

        raise HTTPException(status_code=400, detail=f"Unsupported file type: .{ext}")

    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=422, detail=str(exc))


@app.get("/health")
def health():
    return {"status": "ok"}
