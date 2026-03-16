# -*- coding: utf-8 -*-
"""
Project: サンレモ成約捕捉
File: store_documents.py
Description: 
    1. 作業フォルダの作成
    2. 指定フォルダ内の証跡書類の振分け、およびフォルダ作成

Copyright (c) 2026 SCSK ServiceWare Corporation.
All rights reserved.
"""

import pandas as pd
import xlwings as xw
import datetime
import shutil
import pathlib

from utils import (
    # 設定値
    set_config,
    # 独自エラー
    AppError, ZeroDataError, NotSelectError, ReferencePathError, NoSheetError,
    # 共通関数
    get_pw, select_folder, select_file, search_file, selectfile_to_df, serial_filepath, sanitize_filename, get_base_dir,
    create_name, create_filename
)


# ==== 日付文字列生成 ====
YYYYMMDD = datetime.date.today().strftime("%Y%m%d")
YYYYMM = datetime.date.today().strftime("%Y%m")

config = set_config()


# ============================================
# フォルダーの作成
# ============================================
ROOT_FOLDER = pathlib.Path(config['ルートフォルダ'])
BASE_FOLDER = ROOT_FOLDER / '07_クライアント提出物' 
STORE_FOLDER = BASE_FOLDER / '■振分前書類格納'
PENDING_FOLDER = BASE_FOLDER / '00_未振分フォルダ'
WORKING_FOLDER = BASE_FOLDER / '01_格納フォルダ'
CSV_FOLDER = BASE_FOLDER / '02_CSV'
RESULT_FOLDER = BASE_FOLDER / '99_処理結果'

def make_folders():
    BASE_FOLDER.mkdir(parents=True, exist_ok=True)
    STORE_FOLDER.mkdir(parents=True, exist_ok=True)
    PENDING_FOLDER.mkdir(exist_ok=True)
    WORKING_FOLDER.mkdir(exist_ok=True)
    CSV_FOLDER.mkdir(exist_ok=True)
    RESULT_FOLDER.mkdir(exist_ok=True)

# ============================================
# ファイルの振分け
# ============================================
def store_documents():

    result = []
    
    # 反響リストを取得
    hankyo_df = selectfile_to_df("反響一覧.csv", CSV_FOLDER)

    # 格納フォルダからパスを取得
    files = [f for f in STORE_FOLDER.glob("*") if f.is_file() and not f.name.startswith('~$')]
    if len(files) == 0: 
        raise ZeroDataError(f'証跡ファイルが格納されていません')
    
    # 判定と格納
    for file in files:
        
        report = {
            '振分先CD': '',
            '依頼番号': '',
            'ドキュメント名': str(file.name),
            '振分結果': '',
            'ファイルパス': '',
        }

        # ファイル名から依頼番号取得
        file_name = file.stem
        matched_rows = hankyo_df[hankyo_df.apply(lambda x: x['依頼番号'] in file_name, axis=1)]

        # 企業情報などの取得
        if not matched_rows.empty:
            row = matched_rows.iloc[0]
            furiwakesaki_code = row['振分先コード']
            kigyo_name = row['企業名']
            furiwakesaki_name = row['振分先名']
            irai_no = row['依頼番号']

            report['振分先CD'] = furiwakesaki_code
            report['依頼番号'] = irai_no

            # --- 1. 振分先フォルダの確定 ---
            furi_folders = list(WORKING_FOLDER.glob(f"*{furiwakesaki_code}*"))
            if len(furi_folders) == 1:
                furi_folder = furi_folders[0]
            else:
                if furiwakesaki_name:
                    furi_folder = WORKING_FOLDER / f'{furiwakesaki_code}_{kigyo_name}_{furiwakesaki_name}'
                else:
                    furi_folder = WORKING_FOLDER / f'{furiwakesaki_code}_{kigyo_name}'                    
                furi_folder.mkdir(parents=True, exist_ok=True)

            # --- 2. 依頼番号フォルダの確定 ---
            irai_folders = list(furi_folder.glob(f"*{irai_no}*"))
            if len(irai_folders) == 1:
                irai_folder = irai_folders[0]
            else:
                irai_folder = furi_folder / f'{irai_no}'
                # 成約・着工サブフォルダも作成
                (irai_folder / '成約').mkdir(parents=True, exist_ok=True)
                (irai_folder / '着工').mkdir(parents=True, exist_ok=True)

            # --- 3. 移動 ---
            dest_path = irai_folder / file.name
            if dest_path.exists():
                # 同名があれば、ファイル名の末尾に時刻などを付けて回避する
                now_time = datetime.datetime.now().strftime("%H%M%S")
                dest_path = irai_folder / f"{file.stem}_{now_time}{file.suffix}"

            shutil.move(file, dest_path)

            report['振分結果'] = "成功"
            report['ファイルパス'] = str(dest_path)


        else:
            # 依頼番号が見つからなかった場合
            dest_path = PENDING_FOLDER / file.name
            PENDING_FOLDER.mkdir(exist_ok=True)
            shutil.move(file, dest_path)

            report['振分結果'] = "失敗"
            report['ファイルパス'] = str(dest_path)
            
        result.append(report)

    result_df = pd.DataFrame(result)

    # テンプレートファイル取得
    template_filepath = pathlib.Path(config['ファイルパス']['ドキュメント振り分け結果FMT'])

    # 保存ファイル名(すでにある場合はリネームして新規で作成)
    base_name = f"{YYYYMMDD}_ドキュメント振り分け結果"
    save_filepath = serial_filepath(RESULT_FOLDER, base_name, '.xlsx')
    shutil.copy(template_filepath, save_filepath)

    # pandasの結果を流し込む
    with xw.App(visible=False) as app:
        wb = xw.Book((save_filepath))
        ws = wb.sheets[0]  # 1番左のシートを指定
        
        ws.range("A:B").number_format = '@'
        ws.range('A2').options(index=False, header=False).value = result_df      

        # 保存
        wb.save()
        wb.close()
    
    print(f' >> {save_filepath} を作成')

    # データ蓄積
    integrated_filepath = RESULT_FOLDER / 'ALL_ドキュメント振り分け結果.xlsx'
    if not integrated_filepath.exists():
        shutil.copy(save_filepath, integrated_filepath)
        
        with xw.App(visible=False) as app:
            wb = xw.Book((integrated_filepath))
            ws = wb.sheets[0]  # 1番左のシートを指定
            
            ws.range("A:B").number_format = "@"
            ws.range("F:F").number_format = "yyyy/mm/dd hh:mm:ss"

            ws.range('F1').value = '振分日時'
            last_row = ws.range("C" + str(ws.cells.last_cell.row)).end("up").row
            ws.range(f'F2:F{last_row}').value = datetime.datetime.now()

            # 保存
            wb.save()
            wb.close()

    else:
        result_df['振分日時'] = datetime.datetime.now()

        with xw.App(visible=False) as app:
            wb = xw.Book((integrated_filepath))
            ws = wb.sheets[0]  # 1番左のシートを指定
            
            ws.range("A:B").number_format = "@"
            ws.range("F:F").number_format = "yyyy/mm/dd hh:mm:ss"

            last_row = ws.range("C" + str(ws.cells.last_cell.row)).end('up').row
            ws.range(f"A{last_row + 1}").options(index=False, header=False).value = result_df

            # 保存
            try:
                wb.save()
            except Exception:
                tmp_save_filepath = serial_filepath(RESULT_FOLDER, 'ALL_ドキュメント振り分け結果', '.xlsx')
                wb.save(tmp_save_filepath)
            wb.close()
            
# --- 実行例 ---
if __name__ == "__main__":
        
    try:
        make_folders()

        # ==== 証跡ファイルの振分け ====
        print(f'[{datetime.datetime.now() :%Y-%m-%d %H:%M:%S}] 処理開始')
        store_documents()
        print(f'[{datetime.datetime.now() :%Y-%m-%d %H:%M:%S}] 処理終了')


    except KeyboardInterrupt:
        print('\n⛔ 処理を中断しました')    
    except Exception as e:
        err_name = type(e).__name__
        err_msg = str(e)
        print(f"\n⚠️ エラーが発生しました \nエラー名: {err_name}\n詳細: {err_msg}")
        import traceback
        traceback.print_exc()
