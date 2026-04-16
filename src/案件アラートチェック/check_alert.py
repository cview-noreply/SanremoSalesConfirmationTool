"""
Project: サンレモ成約捕捉
File: check_alert.py
Description: 
    1. 案件アラートのチェック

Copyright (c) 2026 SCSK ServiceWare Corporation.
All rights reserved.
"""

import pandas as pd
import numpy as np
import xlwings as xw
import tkinter as tk
from tkinter import filedialog
import datetime
import ctypes
import sys
from pathlib import Path
import shutil

pd.set_option('future.no_silent_downcasting', True)

MB_SYSTEMMODAL = 0x1000
YYYYMMDD = datetime.datetime.now().strftime('%Y%m%d')


# --- 1. ダイアログでファイルを取得 ---
def select_file(title: str = 'CL向け案件アラート表一覧を選択してください',
                file_types: list = [("すべてのファイル", "*.*")],
                initial_dir: Path = None) -> Path:
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=title, filetypes=file_types, initialdir=initial_dir)
    root.destroy()
    if not file_path:
        ctypes.windll.user32.MessageBoxW(0, 'ファイルが選択されませんでした。', "通知", 0x40 | MB_SYSTEMMODAL)
        sys.exit()
    return Path(file_path)


# --- 2. ダウンロードフォルダから指定ワードでファイルを取得(dfで返す, 選択したファイルはsave_folderに格納) ---
def selectfile_to_df(filename: str, save_folder: Path = None, prefix_date: str = YYYYMMDD) -> pd.DataFrame:
    """dfはfillna('') | filenameに拡張子必要"""

    print(f'{filename}を選択してください')

    dl_folder = Path.home() / "Downloads"

    if filename.endswith('.csv'):
        filepath = select_file(file_types=[('CSVファイル', '*.csv')], initial_dir=str(dl_folder))
        df = pd.read_csv(filepath, encoding='CP932', dtype=str).fillna('')
    elif filename.endswith('.xlsx'):
        filepath = select_file(file_types=[('EXCELファイル', '*.xlsx')], initial_dir=str(dl_folder))
        df = pd.read_excel(filepath, dtype=str).fillna('')
    else:
        print("対応していないファイル形式です")
        return None

    # ★高速化: applymap → map (pandas 2.1+推奨) ＋ str列のみ対象に絞る
    str_cols = df.select_dtypes(include='object').columns
    df[str_cols] = df[str_cols].apply(lambda col: col.str.strip())

    if save_folder:
        dest_path = save_folder / f'{YYYYMMDD}_{filename}'
        dest_path.parent.mkdir(parents=True, exist_ok=True)
        if not dest_path.exists():
            shutil.copy(filepath, dest_path)

    return df


# ====== メイン処理 ======
def check_alert_():

    df = selectfile_to_df('CL向け案件アラート表一覧.csv')
    if len(df) == 0:
        ctypes.windll.user32.MessageBoxW(0, 'データが0件です', "通知", 0x40 | MB_SYSTEMMODAL)

    num_origin = len(df.columns) + 2

    # ★高速化: 日付変換を一括処理
    date_col_map = {
        '資料請求日':         '資料請求日',
        '初回来場日':         '初回来場日',
        '契約予定日':         '契約予定日',
        '着工予定日':         '着工予定日（成約確認書）　※延期の申告を受けた場合は上書き',
        '最終更新日':         '最終更新日（クライアント）',
        'お客様来場日':       'お客様来場日（来場キャンペーン応募情報）',
        'お客様来場聞き取り日': 'お客様来場聞き取り日',
        'お客様成約聞き取り日': 'お客様成約聞き取り日',
    }
    for new_col, src_col in date_col_map.items():
        df[new_col] = pd.to_datetime(df[src_col], errors='coerce').dt.normalize()

    today = pd.Timestamp.now().normalize()

    # ★高速化: よく使う日付境界値を事前計算
    today_minus_30  = today - pd.Timedelta(days=30)
    today_plus_7    = today + pd.Timedelta(days=6)
    today_minus_70  = today - pd.Timedelta(days=70)
    month_start     = today.replace(day=1)
    prev_month_start = month_start - pd.DateOffset(months=1)
    month_end       = today + pd.offsets.MonthEnd(0)

    # ==== 各フィールドの条件 ====

    df['反響種別_来場'] = df['反響種別'].isin(['モデルハウス訪問予約', '出張訪問予約', 'イベント予約（営業所）', 'イベント予約（企業）'])

    df['資料請求日_30日前'] = df['資料請求日'].between(today_minus_30, today)

    df['ご商談状況_契約済み']        = df['ご商談状況'] == '契約済み'
    df['ご商談状況_not契約済み']     = ~df['ご商談状況_契約済み']
    df['ご商談状況_見積・仮契約']    = df['ご商談状況'].isin(['詳細プラン見積り', '仮契約・設計契約'])
    df['ご商談状況_資料送付済み']    = df['ご商談状況'] == '資料送付済み'
    df['ご商談状況_来場済み']        = df['ご商談状況'].isin(['来場・商談', '概算プラン見積り', '詳細プラン見積り', '仮契約・設計契約'])
    df['ご商談状況_来場済み_土地探し'] = df['ご商談状況'].isin(['来場・商談', '概算プラン見積り', '詳細プラン見積り', '仮契約・設計契約', '土地探し中'])

    # ★高速化: isna() で直接判定（regex replace 不要）
    df['初回来場日_null']   = df['初回来場日'].isna()
    df['契約予定日_null']   = df['契約予定日'].isna()

    df['契約予定日_7日以内']  = df['契約予定日'].between(today, today_plus_7)
    df['契約予定日_前日前月'] = (df['契約予定日'] < today)        & (df['契約予定日'] >= prev_month_start)
    df['契約予定日_前々月']   = df['契約予定日'] < prev_month_start
    df['契約予定日_当月']     = df['契約予定日'].between(month_start, month_end)

    df['着工予定日_前日前月'] = (df['着工予定日'] < today)        & (df['着工予定日'] >= prev_month_start)
    df['着工予定日_前々月']   = df['着工予定日'] < prev_month_start
    df['着工予定日_当月']     = df['着工予定日'].between(month_start, month_end)

    df['最終更新日_30_70日前'] = df['最終更新日'].between(today_minus_70, today_minus_30)

    df['お客様来場日_notnull']       = df['お客様来場日'].notna()
    df['お客様来場聞き取り日_notnull'] = df['お客様来場聞き取り日'].notna()
    df['お客様成約聞き取り日_notnull'] = df['お客様成約聞き取り日'].notna()

    df['成約聞き取り経路_アンケート'] = df['成約聞き取り経路'].isin(['架電およびアンケート', '成約アンケート'])
    df['成約聞き取り経路_架電']       = df['成約聞き取り経路'] == '架電ヒアリング'

    df['成約確認書初回提出日_null']   = df['成約確認書初回提出日'].replace('', np.nan).isna()
    df['成約確認書初回提出日_notnull'] = ~df['成約確認書初回提出日_null']

    df['着工確認書初回提出日_null']   = df['着工確認書初回提出日'].replace('', np.nan).isna()
    df['着工確認書初回提出日_notnull'] = ~df['着工確認書初回提出日_null']

    df['成約時不備チェック_確認中'] = df['成約時不備チェック'] == '不備確認中'
    df['着工時不備チェック_確認中'] = df['着工時不備チェック'] == '不備確認中'


    # ==== PUSH区分の付与 ====

    df['区分_1']  = df['成約時不備チェック_確認中'] & df['着工時不備チェック_確認中']
    df['区分_2']  = df['着工時不備チェック_確認中']
    df['区分_3']  = df['成約時不備チェック_確認中']

    df['区分_4']  = df['着工予定日_前々月']   & df['着工確認書初回提出日_null']
    df['区分_5']  = df['着工予定日_前日前月'] & df['着工確認書初回提出日_null']

    df['区分_6']  = df['お客様成約聞き取り日_notnull'] & df['成約聞き取り経路_アンケート'] & df['成約確認書初回提出日_null']
    df['区分_7']  = df['契約予定日_前々月']   & df['成約確認書初回提出日_null']
    df['区分_8']  = df['契約予定日_前日前月'] & df['成約確認書初回提出日_null']
    df['区分_9']  = df['ご商談状況_契約済み'] & df['成約確認書初回提出日_null']
    df['区分_10'] = df['契約予定日_7日以内']  & df['成約確認書初回提出日_null']
    df['区分_11'] = df['契約予定日_当月']      & df['成約確認書初回提出日_null']

    df['区分_12'] = df['着工予定日_当月'] & df['着工確認書初回提出日_null']

    df['区分_13'] = df['ご商談状況_not契約済み'] & df['成約確認書初回提出日_notnull']

    df['区分_14'] = df['ご商談状況_見積・仮契約'] & df['初回来場日_null'] & df['契約予定日_null'] & df['成約確認書初回提出日_null'] & df['着工確認書初回提出日_null']
    df['区分_15'] = df['ご商談状況_見積・仮契約'] & df['契約予定日_null'] & df['成約確認書初回提出日_null'] & df['着工確認書初回提出日_null']

    df['区分_16a'] = df['ご商談状況_資料送付済み']  & df['お客様来場日_notnull'] & df['成約確認書初回提出日_null'] & df['着工確認書初回提出日_null']
    df['区分_16b'] = df['ご商談状況_not契約済み']   & df['お客様来場日_notnull'] & df['初回来場日_null'] & df['成約確認書初回提出日_null'] & df['着工確認書初回提出日_null']

    df['区分_17a'] = df['ご商談状況_資料送付済み']  & df['お客様来場聞き取り日_notnull'] & df['成約確認書初回提出日_null'] & df['着工確認書初回提出日_null']
    df['区分_17b'] = df['ご商談状況_not契約済み']   & df['初回来場日_null'] & df['お客様来場聞き取り日_notnull'] & df['成約確認書初回提出日_null'] & df['着工確認書初回提出日_null']

    df['区分_18'] = df['ご商談状況_来場済み'] & df['初回来場日_null'] & df['成約確認書初回提出日_null'] & df['着工確認書初回提出日_null']
    df['区分_19'] = df['お客様成約聞き取り日_notnull'] & df['成約聞き取り経路_架電'] & df['成約確認書初回提出日_null'] & df['着工確認書初回提出日_null']

    df['区分_20a'] = df['反響種別_来場'] & df['資料請求日_30日前'] & df['ご商談状況_資料送付済み']  & df['成約確認書初回提出日_null'] & df['着工確認書初回提出日_null']
    df['区分_20b'] = df['反響種別_来場'] & df['資料請求日_30日前'] & df['初回来場日_null'] & df['ご商談状況_not契約済み'] & df['成約確認書初回提出日_null'] & df['着工確認書初回提出日_null']

    df['区分_21'] = df['ご商談状況_来場済み_土地探し'] & df['最終更新日_30_70日前'] & df['成約確認書初回提出日_null'] & df['着工確認書初回提出日_null']


    # ==== 一番左の区分を取得 ====
    kubun_cols = [c for c in df.columns if c.startswith('区分_')]

    # ★高速化: any() の結果を先にキャッシュして idxmax と重複走査を避ける
    kubun_any = df[kubun_cols].any(axis=1)
    df['PUSH対象区分_チェック'] = df[kubun_cols].idxmax(axis=1).where(kubun_any, '')

    # 区分名から管理コードへの置換
    kubun_map = {
        '区分_1':  'A 1',  '区分_2':  'A 2',  '区分_3':  'A 3',
        '区分_4':  'A 4',  '区分_5':  'A 5',  '区分_6':  'A 6',
        '区分_7':  'A 7',  '区分_8':  'A 8',  '区分_9':  'A 9',
        '区分_10': 'A 10', '区分_11': 'A 11', '区分_12': 'A 12',
        '区分_13': 'A 13', '区分_14': 'A 14', '区分_15': 'A 15',
        '区分_16a': 'A 16', '区分_16b': 'A 16',
        '区分_17a': 'A 17', '区分_17b': 'A 17',
        '区分_18': 'A 18',
        '区分_19': 'B 19', '区分_20a': 'B 20', '区分_20b': 'B 20',
        '区分_21': 'B 21',
    }
    df['PUSH対象区分_チェック'] = df['PUSH対象区分_チェック'].replace(kubun_map)

    # 一致チェック
    # ★修正: 元コードは replace の戻り値を代入していないバグがあるため修正
    check_series = df['PUSH対象区分_チェック'].replace(r'^\s*$', np.nan, regex=True).fillna('')
    push_series  = df['PUSH対象区分'].replace(r'^\s*$', np.nan, regex=True).fillna('')
    df['正誤判定'] = (check_series == push_series).replace({True: "OK", False: "NG"})

    counts = str(df['正誤判定'].value_counts()).replace('Name: count, dtype: int64', '')

    #  True/False → ●/"" の置換をブール列のみに絞る（全列走査を回避）
    bool_cols = df.select_dtypes(include='bool').columns
    df[bool_cols] = df[bool_cols].replace({True: "●", False: ""})

    # 列を一括削除
    drop_cols = ['着工予定日', '最終更新日', 'お客様来場日']
    df = df.drop(columns=drop_cols, errors='ignore')

    # 列の並べ替え
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    first_cols = ['正誤判定', 'PUSH対象区分', 'ご確認ご依頼事項', '概要', '振分先コード', '企業名', '振分先名']
    other_cols = [c for c in df.columns if c not in first_cols]
    df = df[first_cols + other_cols]

    # 日付型の列を「YYYY/MM/DD」形式の文字列に変換
    for col in df.select_dtypes(include=['datetime']).columns:
        df[col] = df[col].dt.strftime('%Y/%m/%d')

    # インデックス1始まり
    df.index = df.index + 1

    # 保存
    try:
        save_filename = f'{YYYYMMDD}_案件アラートチェックシート.xlsx'
        df.to_excel(save_filename)
    except Exception:
        save_filename = f'{YYYYMMDD}_案件アラートチェックシート_{datetime.datetime.now().strftime("%H%M%S")}.xlsx'
        df.to_excel(save_filename)

    with xw.App(visible=False) as app:
        wb = xw.Book(save_filename)
        sht = wb.sheets[0]

        table_range = sht.range("A1").expand()
        total_cols = table_range.columns.count
        total_rows = table_range.rows.count

        # --- 1. 全体の基本設定 ---
        table_range.api.Font.Name = "メイリオ"
        table_range.api.Font.Size = 10
        table_range.api.VerticalAlignment = -4108  # xlCenter

        # --- 2. ヘッダーの色分け ---
        sht.range((1, 1), (1, num_origin)).color = (204, 255, 204)

        if total_cols > num_origin:
            sht.range((1, num_origin + 1), (1, total_cols)).color = (204, 229, 255)
            sht.range((1, num_origin + 1), (total_rows, total_cols)).api.HorizontalAlignment = -4108

        header_all = sht.range((1, 1), (1, total_cols))
        header_all.api.Font.Bold = True
        header_all.api.Font.Color = 0x000000

        # --- 3. 罫線の付与 ---
        for i in range(7, 13):
            border = table_range.api.Borders(i)
            border.LineStyle = 1
            border.Weight = 2

        # --- 4. 列幅調整・保存 ---
        table_range.autofit()
        wb.save()
        wb.close()

    ctypes.windll.user32.MessageBoxW(0, f'{save_filename}の作成が完了しました。 {counts}', "通知", 0x40 | MB_SYSTEMMODAL)


if __name__ == '__main__':
    check_alert_()