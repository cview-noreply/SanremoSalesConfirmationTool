
import pandas as pd
import numpy as np
import random
from datetime import date, timedelta
from pathlib import Path


# 書き換え
file_path = r"...テストデータ作成用.xlsx"

test_data_folder = Path('テスト用ファイル')
dl_folder = Path.home() / "Downloads"

xls = pd.ExcelFile(file_path)
sheet_names = xls.sheet_names

for sheet_name in sheet_names:
    df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)

    if 'csv' in sheet_name:
        df.to_csv(test_data_folder / sheet_name, encoding='CP932', index=False)
        df.to_csv(dl_folder / sheet_name, encoding='CP932', index=False)
    
    elif 'xlsx' in sheet_name:
        df.to_excel(test_data_folder / sheet_name, index=False)
        df.to_excel(dl_folder / sheet_name, index=False)

