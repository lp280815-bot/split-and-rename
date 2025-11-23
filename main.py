# main.py
import io
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
import openpyxl

from giyul_logic import process_workbook

app = FastAPI(title="Giyul Chovot Processor")


@app.post("/process")
async def process_excel(file: UploadFile = File(...)):
    """
    Принимает Excel-файл, запускает все логики 1–7
    и возвращает обработанный Excel.
    """
    if not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="Нужен файл Excel (.xlsx / .xlsm)")

    contents = await file.read()

    try:
        wb = openpyxl.load_workbook(io.BytesIO(contents))
        wb = process_workbook(wb)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка обработки файла: {e}")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    out_name = file.filename.rsplit(".", 1)[0] + "_processed.xlsx"

    headers = {
        "Content-Disposition": f'attachment; filename="{out_name}"'
    }

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )
