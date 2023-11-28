import os
from pathlib import Path

import xlwings as xl
from dotenv import load_dotenv


def get_model_path():
    load_dotenv()
    model_file_path = Path(os.getenv("EXCEL_MODELS_PATH"))
    model_file_name = os.getenv("MODEL_NAME", "not_found")

    the_path: Path = model_file_path / model_file_name

    return the_path


def try_with_xlwings(num1: float, num2: float):
    sheet_idx = 0

    with xl.App(visible=False) as app:
        wb = app.books.open(get_model_path())
        wb.app.calculation = "manual"

        wb.sheets[sheet_idx]["B1"].value = num1
        wb.sheets[sheet_idx]["B2"].value = num2

        wb.app.calculate()
        the_result = wb.sheets["demo_model"]["B3"].value

        # don't save
        # wb.save()
        wb.close()

    return the_result
