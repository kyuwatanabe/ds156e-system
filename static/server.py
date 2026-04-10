import os
import io
import json
import pandas as pd
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import Response, JSONResponse
from fastapi.staticfiles import StaticFiles
import anthropic
from pypdf import PdfReader, PdfWriter

app = FastAPI()

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
PDF_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "ds156_e.pdf")

DEFAULT_SYSTEM_PROMPT = """\
# DS-156E 財務情報抽出システムプロンプト
# ========================================
# このプロンプトはExcelの財務諸表からDS-156Eフォームに必要な情報を抽出します。
# 財務諸表のスタイルが異なる場合は、以下のルールを修正してください。
# 「#」で始まる行はコメントです（Claudeへの指示には含まれません）。
#
# 【対応している財務諸表スタイル】
# - 英語の連結損益計算書＋バランスシート（YTDシート形式）
# - 日本語の貸借対照表＋損益計算書
# - 単体・連結どちらも対応
#
# 【Total Assets（総資産）の抽出ルール】
# - "Total Assets", "総資産", "資産合計", "資産の部 合計" などのラベルを探す
# - バランスシートの資産側の最終合計値を使用する
#
# 【Total Liabilities（総負債）の抽出ルール】
# - "Total Liabilities", "負債合計", "負債の部 合計" などを探す
#
# 【Owner's Equity（純資産）の抽出ルール】
# - "Total Stockholder's Equity", "純資産合計", "自己資本合計" などを探す
# - マイナスの場合はそのまま負の値として返す
#
# 【Operating Income Before Taxes（税引前利益）の抽出ルール】
# - "Pre Tax Profit", "税引前当期純利益", "経常利益" などを探す
#
# 【Operating Income After Taxes（税引後利益）の抽出ルール】
# - "Net Income", "当期純利益" などを探す
#
# 【Inventory（棚卸資産）の抽出ルール】
# - "Inventory", "棚卸資産", "商品及び製品" などを探す
# - グロス（引当金控除前）の値を使用する
# - 在庫引当金（Inventory Reserve）は含めない
#
# 【Equipment（設備）の抽出ルール】
# - "Plant & Equipment", "有形固定資産", "機械装置" などを探す
# - グロス（減価償却累計額控除前）の値を使用する
# - 減価償却累計額（Accumulated Depreciation）は含めない
#
# 【Cash（現金）の抽出ルール】
# - "Cash", "現金及び預金", "Petty Cash" + "Cash-Bank" の合計
#
# 【Premises（不動産・土地建物）の抽出ルール】
# - "Land", "Building", "土地", "建物" などを探す
# - バランスシートに明示的な項目がなければ0とする
#
# 【通貨・単位の処理】
# - 金額はUSDで返す
# - 日本円の場合は currency を "JPY" にする
# - 単位が「千円」「百万円」の場合は実額に換算する
#
# 【複数シートがある場合】
# - YTD（Year to Date）または年間合計シートを優先する
# - 月次シート（P12など）は使用しない

You are a financial data extraction assistant for DS-156E visa applications.
Extract the following financial figures from the Excel data provided.
Return ONLY a valid JSON object. No explanation, no markdown, no code blocks.

Required JSON format:
{
  "year": "2025",
  "total_assets": 61913640.00,
  "total_liabilities": 89614070.00,
  "owners_equity": -27700430.00,
  "income_before_tax": 259440.00,
  "income_after_tax": 194540.00,
  "cash": 2337470.00,
  "inventory": 5617918.00,
  "equipment": 6546970.00,
  "premises": 0.00,
  "currency": "USD",
  "notes": "抽出時の特記事項があれば記載"
}
"""


def parse_excel_to_text(file_bytes: bytes) -> str:
    """ExcelファイルをテキストとしてClaudeに渡せる形式に変換"""
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    result = []
    for sheet_name in xl.sheet_names:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
        result.append(f"=== Sheet: {sheet_name} ===")
        result.append(df.to_string(na_rep=""))
    return "\n\n".join(result)


def extract_system_prompt(raw_prompt: str) -> str:
    """コメント行（#で始まる行）を除去してClaudeへ渡すプロンプトを生成"""
    lines = raw_prompt.split("\n")
    filtered = [line for line in lines if not line.strip().startswith("#")]
    return "\n".join(filtered).strip()


def format_usd(value: float) -> str:
    """数値をDS-156E用のドル表示にフォーマット"""
    if value < 0:
        return f"-${abs(value):,.0f}"
    return f"${value:,.0f}"


def fill_pdf(data: dict) -> bytes:
    """抽出した財務データでDS-156E PDFを埋める"""
    reader = PdfReader(PDF_TEMPLATE_PATH)
    writer = PdfWriter()
    writer.append(reader)

    assets = data.get("total_assets", 0)
    equity = data.get("owners_equity", 0)

    # #11: Fair Market Value の計算
    # Owner's Equityがマイナスの場合はTotal Assetsの3倍、それ以外はEquityの3倍
    if equity < 0:
        fair_market_value = assets * 3
    else:
        fair_market_value = equity * 3

    # #12: Other = Total Assets - Cash - Inventory - Equipment - Premises
    cash = data.get("cash", 0)
    inventory = data.get("inventory", 0)
    equipment = data.get("equipment", 0)
    premises = data.get("premises", 0)
    other = assets - cash - inventory - equipment - premises

    fields = {
        # #9 Financial Statement
        "StateYr": str(data.get("year", "")),
        "FinCY": "/Yes",
        "Assets": format_usd(assets),
        "Liabil": format_usd(data.get("total_liabilities", 0)),
        "Equity": format_usd(equity),
        "BefTax": format_usd(data.get("income_before_tax", 0)),
        "AftTax": format_usd(data.get("income_after_tax", 0)),
        # #11 Type of Investment
        "ExBus": "/Yes",
        "EBValue": format_usd(fair_market_value),
        # #12 Total Investment from Abroad (Cumulative)
        "CashCum": format_usd(cash),
        "InvCum": format_usd(inventory),
        "EqpCum": format_usd(equipment),
        "PreCum": format_usd(premises),
        "OthCum": format_usd(other),
        "TotCum": format_usd(assets),
    }

    for page in writer.pages:
        if "/Annots" in page:
            for annot in page["/Annots"]:
                annot_obj = annot.get_object()
                field_name = annot_obj.get("/T")
                if field_name and field_name in fields:
                    value = fields[field_name]
                    if value == "/Yes":
                        annot_obj.update({
                            "/V": "/Yes",
                            "/AS": "/Yes",
                        })
                    else:
                        annot_obj.update({"/V": value, "/AP": {}})

    output = io.BytesIO()
    writer.write(output)
    return output.getvalue()


@app.get("/default-prompt")
def get_default_prompt():
    return JSONResponse({"prompt": DEFAULT_SYSTEM_PROMPT})


@app.post("/extract")
async def extract_financial_data(
    file: UploadFile = File(...),
    system_prompt: str = Form(DEFAULT_SYSTEM_PROMPT),
):
    """ExcelファイルからClaudeを使って財務データを抽出"""
    if not ANTHROPIC_API_KEY:
        raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY が設定されていません")

    file_bytes = await file.read()
    try:
        excel_text = parse_excel_to_text(file_bytes)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Excelファイルの読み込みエラー: {str(e)}")

    clean_prompt = extract_system_prompt(system_prompt)

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1000,
        system=clean_prompt,
        messages=[
            {
                "role": "user",
                "content": f"以下のExcelデータから財務情報を抽出してください:\n\n{excel_text}"
            }
        ],
    )

    raw = message.content[0].text.strip()
    # コードブロックがあれば除去
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip()

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        raise HTTPException(status_code=500, detail=f"Claudeの応答をJSONとして解析できませんでした: {raw}")

    return JSONResponse(data)


@app.post("/generate-pdf")
async def generate_pdf(data: dict):
    """財務データからDS-156E PDFを生成"""
    try:
        pdf_bytes = fill_pdf(data)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PDF生成エラー: {str(e)}")

    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={"Content-Disposition": "attachment; filename=ds156e_filled.pdf"},
    )


app.mount("/", StaticFiles(directory="static", html=True), name="static")
