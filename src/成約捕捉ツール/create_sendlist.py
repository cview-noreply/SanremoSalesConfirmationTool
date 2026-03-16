# -*- coding: utf-8 -*-
"""
Project: サンレモ成約捕捉
File: create_sendlist.py
Description: 
    1. 案件管理シートの一括送付
    2. 成約確認書・着工証跡引き取り便送付
    3. 個別着工証跡引き取り便送付
    4. 案件管理シートリマインド送付
    5. HONEY管理クライアントへの案件進捗メール一括送付
    6. 案件アラートファイル一括送付

Copyright (c) 2026 SCSK ServiceWare Corporation.
All rights reserved.
"""

import pandas as pd
import xlwings as xw
import datetime
from pathlib import Path
import traceback
import tkinter as tk
from tkinter import messagebox

from utils import (
    # 設定値
    set_config,
    # 独自エラー
    AppError, ZeroDataError, NotSelectError, ReferencePathError, NoSheetError,
    # 共通関数
    get_pw, select_folder, select_file, search_file, selectfile_to_df, serial_filepath, sanitize_filename, get_base_dir,
    create_name, create_filename
)


# ============================================
# 各種変数・関数
# ============================================
# 日付文字列生成
YYYYMMDD = datetime.date.today().strftime("%Y%m%d")
YYYYMM = datetime.date.today().strftime("%Y%m")
YYMM = datetime.date.today().strftime("%y%m")


BASE_DIR = get_base_dir()
config = set_config()

# ============================================
# フォルダーの作成
# ============================================
ROOT_FOLDER = Path(config['ルートフォルダ'])
BASE_FOLDER = ROOT_FOLDER / '03_メール送付・引き取り'

def make_folders(folders_list:list):
    for folder in folders_list:
        folder.mkdir(exist_ok=True, parents=True)
        
    # 2. メッセージボックスを表示
    # メインウィンドウを隠すおまじない
    root = tk.Tk()
    root.withdraw()
    # 1. 最前面に設定
    root.attributes('-topmost', True)
    # 2. 強制的にフォーカスを当てる
    root.focus_force()

    messagebox.showinfo("確認", "フォルダを作成しました。\n対象ファイルを格納してください。")

    # 使い終わったら破棄
    root.destroy()


# ==================================================================================================
# 1. 案件管理シートの一括送付
# ==================================================================================================
def create_aks_bulk():

    # 作業フォルダ作成
    BASE_FOLDER_SEC = BASE_FOLDER / f'案件管理シート一括送付用'
    WORKING_FOLDER = BASE_FOLDER_SEC / YYYYMMDD
    INPUT_FOLDER = WORKING_FOLDER / 'インプットデータ'

    
    # CSV格納の猶予与える
    if not WORKING_FOLDER.exists():
        make_folders([WORKING_FOLDER, INPUT_FOLDER])
        print("フォルダを作成しました. CSVを格納後、再度処理を実行してください")
        return


    # ファイルの取得
    send_df = selectfile_to_df("案件管理シート送付先リスト.csv", INPUT_FOLDER)
    person_df = selectfile_to_df("案件管理担当者リスト.csv", INPUT_FOLDER)
    target_df = selectfile_to_df("案件管理シート送付対象クライアント一覧.xlsx", INPUT_FOLDER)

    # --- 1. 結合 --- 
    merged_df1 = pd.merge(
        send_df[['振分先コード', '案件管理主体の振分先コード', '企業名', '振分先名']], 
        person_df[['振分先コード', '担当者部署', '担当者氏名', '担当者メールアドレス']], 
        left_on='案件管理主体の振分先コード', 
        right_on='振分先コード', 
        how='inner'
    )

    # 列整理
    merged_df1 = merged_df1.rename(columns={'振分先コード_x': '振分先コード'})
    merged_df1 = merged_df1.drop(columns=['案件管理主体の振分先コード', '振分先コード_y'])
    
    # 出力
    base_name = f"{YYYYMMDD}_送付先リストx担当者リスト"
    save_filepath = serial_filepath(WORKING_FOLDER, base_name,'.csv')
    merged_df1.to_csv(save_filepath, index=False, encoding="CP932")


    # --- 2. 結合 --- 
    merged_df2 = pd.merge(
        merged_df1, 
        target_df[['振分先コード', 'ファイル名']], 
        left_on='振分先コード', 
        right_on='振分先コード', 
        how='inner'
    )


    # --- 3. 整形 --- 
    def create_name_format(row):
        kigyo = row['企業名']
        furi = row['振分先名']
        busho = row['担当者部署']
        shimei = row['担当者氏名']
        mail = row['担当者メールアドレス']
        
        # 振分先名がある場合だけカッコを付ける
        furi_part = f" ({furi}分)" if furi != "" else ""
        
        # 全体を組み立てる
        return f"{kigyo}{furi_part} {busho} {shimei} <{mail}>"

    # 新しいデータフレームを作成
    output_df = pd.DataFrame()
    output_df['名前'] = merged_df2.apply(create_name_format, axis=1)
    output_df['ファイル名'] = merged_df2['ファイル名']

    # --- 4. 出力 --- 
    base_name = f"{YYYYMMDD}_案件管理シート宛先取込用"
    save_filepath = serial_filepath(WORKING_FOLDER, base_name,'.csv')
    output_df.to_csv(save_filepath, index=False, header=False, encoding="CP932")

    print(f' >> {str(save_filepath.name)} を作成')

# ==================================================================================================
# 2. 成約確認書・着工証跡引き取り便送付
# ==================================================================================================
def create_doc_pickup():


    # 作業フォルダ作成
    BASE_FOLDER_SEC = BASE_FOLDER / f'成約確認書・着工証跡引き取り便送付'
    WORKING_FOLDER = BASE_FOLDER_SEC / YYYYMMDD
    INPUT_FOLDER = WORKING_FOLDER / 'インプットデータ'
    
    # CSV格納の猶予与える
    if not WORKING_FOLDER.exists():
        make_folders([WORKING_FOLDER, INPUT_FOLDER])
        print("フォルダを作成しました. CSVを格納後、再度処理を実行してください")
        return

    # ファイルの取得
    send_df = selectfile_to_df("引き取り便送付先リスト.csv", INPUT_FOLDER)
    seiyaku_df = selectfile_to_df("成約報告担当者リスト.csv", INPUT_FOLDER)
    chakko_df = selectfile_to_df("着工証跡担当者リスト.csv", INPUT_FOLDER)


    # --- 1. 結合 ---
    merged_df1 = pd.merge(
        send_df[['振分先コード', '成約報告主体の振分先コード', '企業名', '振分先名']], 
        seiyaku_df[['振分先コード', '担当者部署', '担当者氏名', '担当者メールアドレス']], 
        left_on='成約報告主体の振分先コード', 
        right_on='振分先コード', 
        how='inner'
    )
    # 列整理
    merged_df1 = merged_df1.rename(columns={'振分先コード_x': '振分先コード'})
    merged_df1 = merged_df1.drop(columns=['成約報告主体の振分先コード', '振分先コード_y'])

    # 出力
    base_name = f"{YYYYMMDD}_送付先リストx成約担当者リスト"
    save_filepath = serial_filepath(WORKING_FOLDER, base_name,'.csv')
    merged_df1.to_csv(save_filepath, index=False, encoding="CP932")


    # --- 2. 結合 --- 
    merged_df2 = pd.merge(
        send_df[['振分先コード', '着工報告主体の振分先コード', '企業名', '振分先名']], 
        chakko_df[['振分先コード', '担当者部署', '担当者氏名', '担当者メールアドレス']], 
        left_on='着工報告主体の振分先コード', 
        right_on='振分先コード', 
        how='inner'
    )
    # 列整理
    merged_df2 = merged_df2.rename(columns={'振分先コード_x': '振分先コード'})
    merged_df2 = merged_df2.drop(columns=['着工報告主体の振分先コード', '振分先コード_y'])

    # 出力
    base_name = f"{YYYYMMDD}_送付先リストx着工担当者リスト"
    save_filepath = serial_filepath(WORKING_FOLDER, base_name,'.csv')
    merged_df2.to_csv(save_filepath, index=False, encoding="CP932")


    # --- 3. 縦結合 --- 
    combined_df = pd.concat([merged_df1, merged_df2], axis=0)
    # 列整理
    combined_df = combined_df.drop(columns=['振分先コード'])
    combined_df = combined_df.drop_duplicates()

    # 出力
    base_name = f"{YYYYMMDD}_送付先リストx担当者リスト"
    save_filepath = serial_filepath(WORKING_FOLDER, base_name,'.csv')
    combined_df.to_csv(save_filepath, index=False, encoding="CP932")

    #  --- 4. 整形 --- 
    def create_name_format(row):
        kigyo = row['企業名']
        furi = row['振分先名']
        busho = row['担当者部署']
        shimei = row['担当者氏名']
        
        # 振分先名がある場合だけカッコを付ける
        furi_part = f" ({furi}分)" if furi != "" else ""
        
        # 全体を組み立てる
        return f"{kigyo}{furi_part} {busho} {shimei}"

    # 新しいデータフレームを作成
    output_df = pd.DataFrame()
    output_df['名前'] = combined_df.apply(create_name_format, axis=1)
    output_df['メールアドレス'] = combined_df['担当者メールアドレス']
    output_df['グループ'] = f'{YYMM}_引き取り便'
    output_df['コメント'] = ""
    output_df['属性1'] = ""
    output_df['属性2'] = ""

    # 並び替え
    cols = ['グループ', '名前', 'メールアドレス', 'コメント', '属性1', '属性2']
    output_df = output_df[cols]

    # --- 5. 出力 (1000件ごとに分割) --- 
    chunk_size = 1000
    base_name_template = f"{YYYYMMDD}_引き取り便アドレス帳取込用"

    # 全件数がchunk_size(1000)を超えるかどうかでループ
    for i, start_idx in enumerate(range(0, len(output_df), chunk_size)):
        chunk = output_df[start_idx : start_idx + chunk_size]
        
        # 複数ファイルになる場合はファイル名に連番(_1, _2...)を付与
        if len(output_df) > chunk_size:
            base_name = f"{base_name_template}({i+1})"
        else:
            base_name = base_name_template
            
        save_filepath = serial_filepath(WORKING_FOLDER, base_name, '.csv')
        chunk.to_csv(save_filepath, index=False, encoding="CP932")

        print(f' >> {str(save_filepath.name)} を作成 ({len(chunk)}件)')


# ==================================================================================================
# 3. 個別着工証跡引き取り便送付
# ==================================================================================================
def create_doc_pickup_indv():

    # 作業フォルダ作成
    BASE_FOLDER_SEC = BASE_FOLDER / f'個別着工証跡引き取り便送付'
    WORKING_FOLDER = BASE_FOLDER_SEC / YYYYMMDD
    INPUT_FOLDER = WORKING_FOLDER / 'インプットデータ'
    FILES_FOLDER = WORKING_FOLDER / '着工報告ご依頼ファイル'

    # CSV格納の猶予与える
    if not WORKING_FOLDER.exists():
        make_folders([WORKING_FOLDER, INPUT_FOLDER, FILES_FOLDER])
        print("フォルダを作成しました. CSVを格納後、再度処理を実行してください")
        return


    # ファイルの取得
    person_df = selectfile_to_df("着工報告担当者リスト（特定案件のみ）.csv", INPUT_FOLDER)
    person_df = person_df.rename(columns={'着工予定日（成約確認書）　※延期の申告を受けた場合は上書き':'着工予定日'})

    # 着工報告ご依頼のFMT取得
    template_filepath = config['ファイルパス']['着工報告ご依頼FMT']

    #  --- 整形 --- 
    current_month = datetime.datetime.now().strftime("%Y年%m月")
    def create_name_format(row):
        kigyo = row['企業名']
        furi = row['振分先名']
        busho = row['この案件のみの着工報告担当者部署']
        shimei = row['この案件のみの着工報告担当者氏名']
        mail = row['この案件のみの着工報告担当者メールアドレス']
        
        # 振分先名がある場合だけカッコを付ける
        furi_part = f" ({furi}分)" if furi != "" else ""
        
        # 全体を組み立てる
        return f"{kigyo}{furi_part} {busho} {shimei}<{mail}>"
    
    def create_filename_format(row):
        irai = row['依頼番号']
        kigyo = row['企業名']
        furi = row['振分先名']
        # 振分先名がある場合だけカッコを付ける
        furi_part = f" ({furi}分)" if furi != "" else ""
        
        # 全体を組み立てる
        return f"【SUUMO注文】{current_month}度着工報告ご依頼_{irai}_{kigyo}様{furi_part}.xlsx"

    # person_df自体に列を追加
    person_df['名前'] = person_df.apply(create_name_format, axis=1)
    person_df['ファイル名'] = person_df.apply(create_filename_format, axis=1)

    # --- 出力用データフレームの作成 ---
    output_df = pd.DataFrame()
    output_df['名前'] = person_df['名前']
    output_df['ファイル名'] = person_df['ファイル名']

    # 出力
    base_name = f"{YYYYMMDD}_特定案件の着工報告便宛先取込用"
    save_filepath = serial_filepath(WORKING_FOLDER, base_name, '.csv')
    output_df.to_csv(save_filepath, index=False, header=False, encoding="CP932")
    print(f' >> {str(save_filepath.name)} を作成')
    
    
    # === 着工報告ご依頼ファイルの作成 ===
    print(" -- 着工報告ご依頼ファイル作成開始 -- ")
    irai_df = person_df[['依頼番号', '着工予定日', 'ファイル名']]
    file_count = 0

    # Excelの起動(App)はループの外に出して高速化！
    with xw.App(visible=False) as app:
        
        # 依頼番号ごとにグループ化してループを回す
        for irai_no, group in irai_df.groupby('依頼番号'):

            split_df = group[['依頼番号', '着工予定日']]
            
            # Seriesから文字列（先頭の1件）を取り出す
            filename = group['ファイル名'].iloc[0]
            save_path = FILES_FOLDER / filename

            wb = app.books.open(template_filepath)
            ws = wb.sheets[0]  # 1番左のシートを指定
            
            # 2行目・1列目（A2セル）から値を貼り付け
            ws.range("A:A").number_format = '@'
            ws.range("A2").options(index=False, header=False).value = split_df
            
            # 名前をつけて保存
            wb.save(save_path)
            wb.close()
            file_count += 1
    
    print(f" -- 着工報告ご依頼ファイル作成完了({file_count}件) -- ")


# ==================================================================================================
# 4. 案件管理シートリマインド
# ==================================================================================================
def create_aks_remind():

    # 作業フォルダ作成
    BASE_FOLDER_SEC = BASE_FOLDER / f'案件管理シートリマインド'
    WORKING_FOLDER = BASE_FOLDER_SEC / YYYYMMDD
    INPUT_FOLDER = WORKING_FOLDER / 'インプットデータ'

    # CSV格納の猶予与える
    if not WORKING_FOLDER.exists():
        make_folders([WORKING_FOLDER, INPUT_FOLDER])
        print("フォルダを作成しました. CSVを格納後、再度処理を実行してください")
        return

    # ファイルの取得
    send_df = selectfile_to_df("案件管理シート未提出クライアントリスト.csv", INPUT_FOLDER)
    person_df = selectfile_to_df("案件管理担当者リスト.csv", INPUT_FOLDER)


    # --- 1. 結合 ---
    merged_df1 = pd.merge(
        send_df[['振分先コード', '案件管理主体の振分先コード', '企業名', '振分先名']], 
        person_df[['振分先コード', '担当者部署', '担当者氏名', '担当者メールアドレス']], 
        left_on='案件管理主体の振分先コード', 
        right_on='振分先コード', 
        how='inner'
    )
    # 列整理
    merged_df1 = merged_df1.rename(columns={'振分先コード_x': '振分先コード'})
    merged_df1 = merged_df1.drop(columns=['案件管理主体の振分先コード', '振分先コード_y'])

    # 出力
    base_name = f"{YYYYMMDD}_未提出クライアントリストx担当者リスト"
    save_filepath = serial_filepath(WORKING_FOLDER, base_name,'.csv')
    merged_df1.to_csv(save_filepath, index=False, encoding="CP932")


    #  --- 2 整形 --- 
    def create_name_format(row):
        kigyo = row['企業名']
        furi = row['振分先名']
        busho = row['担当者部署']
        shimei = row['担当者氏名']
        
        # 振分先名がある場合だけカッコを付ける
        furi_part = f" ({furi}分)" if furi != "" else ""
        
        # 全体を組み立てる
        return f"{kigyo}{furi_part} {busho} {shimei}"

    # 新しいデータフレームを作成
    output_df = pd.DataFrame()
    output_df['名前'] = merged_df1.apply(create_name_format, axis=1)
    output_df['メールアドレス'] = merged_df1['担当者メールアドレス']
    output_df['グループ'] = f'{YYMM}_案件管理シート未提出PUSH'
    output_df['コメント'] = ""
    output_df['属性1'] = ""
    output_df['属性2'] = ""

    # 並び替え
    cols = ['グループ', '名前', 'メールアドレス', 'コメント', '属性1', '属性2']
    output_df = output_df[cols]

    # --- 5. 出力 (1000件ごとに分割) --- 
    chunk_size = 1000
    base_name_template = f"{YYYYMMDD}_案件管理シート未提出PUSHアドレス帳取込用"

    # 全件数がchunk_size(1000)を超えるかどうかでループ
    for i, start_idx in enumerate(range(0, len(output_df), chunk_size)):
        chunk = output_df[start_idx : start_idx + chunk_size]
        
        #   (_1, _2...)を付与
        if len(output_df) > chunk_size:
            base_name = f"{base_name_template}({i+1})"
        else:
            base_name = base_name_template
            
        save_filepath = serial_filepath(WORKING_FOLDER, base_name, '.csv')
        chunk.to_csv(save_filepath, index=False, encoding="CP932")

        print(f' >> {str(save_filepath.name)} を作成 ({len(chunk)}件)')



# ==================================================================================================
# 5. HONEY管理クライアントへの案件進捗メール一括送付
# ==================================================================================================
def create_Honey_progress():

    # 作業フォルダ作成
    BASE_FOLDER_SEC = BASE_FOLDER / f'HONEY管理クライアントへの案件進捗メール一括送付'
    WORKING_FOLDER = BASE_FOLDER_SEC / YYYYMMDD
    INPUT_FOLDER = WORKING_FOLDER / 'インプットデータ'

    if not WORKING_FOLDER.exists():
        make_folders([WORKING_FOLDER, INPUT_FOLDER])
        print("フォルダを作成しました. CSVを格納後、再度処理を実行してください")
        return

    # ファイルの取得
    send_df = selectfile_to_df("HONEY管理リマインド送付先リスト.csv", INPUT_FOLDER)
    person_df = selectfile_to_df("案件管理担当者リスト.csv", INPUT_FOLDER)

    # --- 1. 結合 ---
    merged_df1 = pd.merge(
        send_df[['振分先コード', '案件管理主体の振分先コード', '企業名', '振分先名']], 
        person_df[['振分先コード', '担当者部署', '担当者氏名', '担当者メールアドレス']], 
        left_on='案件管理主体の振分先コード', 
        right_on='振分先コード', 
        how='inner'
    )
    # 列整理
    merged_df1 = merged_df1.rename(columns={'振分先コード_x': '振分先コード'})
    merged_df1 = merged_df1.drop(columns=['案件管理主体の振分先コード', '振分先コード_y'])

    # 出力
    base_name = f"{YYYYMMDD}_HONEY管理リマインド送付先リストx担当者リスト"
    save_filepath = serial_filepath(WORKING_FOLDER, base_name,'.csv')
    merged_df1.to_csv(save_filepath, index=False, encoding="CP932")


    #  --- 2 整形 --- 
    def create_name_format(row):
        kigyo = row['企業名']
        furi = row['振分先名']
        busho = row['担当者部署']
        shimei = row['担当者氏名']
        
        # 振分先名がある場合だけカッコを付ける
        furi_part = f" ({furi}分)" if furi != "" else ""
        
        # 全体を組み立てる
        return f"{kigyo}{furi_part} {busho} {shimei}"

    # 新しいデータフレームを作成
    output_df = pd.DataFrame()
    output_df['名前'] = merged_df1.apply(create_name_format, axis=1)
    output_df['メールアドレス'] = merged_df1['担当者メールアドレス']
    output_df['グループ'] = f'{YYMM}_HONEY案件進捗'
    output_df['コメント'] = ""
    output_df['属性1'] = ""
    output_df['属性2'] = ""

    # 並び替え
    cols = ['グループ', '名前', 'メールアドレス', 'コメント', '属性1', '属性2']
    output_df = output_df[cols]


    # --- 5. 出力 (1000件ごとに分割) --- 
    chunk_size = 1000
    base_name_template = f"{YYYYMMDD}_HONEY案件進捗依頼アドレス帳取込用"

    # 全件数がchunk_size(1000)を超えるかどうかでループ
    for i, start_idx in enumerate(range(0, len(output_df), chunk_size)):
        chunk = output_df[start_idx : start_idx + chunk_size]
        
        # 複数ファイルになる場合はファイル名に連番(_1, _2...)を付与
        if len(output_df) > chunk_size:
            base_name = f"{base_name_template}({i+1})"
        else:
            base_name = base_name_template
            
        save_filepath = serial_filepath(WORKING_FOLDER, base_name, '.csv')
        chunk.to_csv(save_filepath, index=False, encoding="CP932")

        print(f' >> {str(save_filepath.name)} を作成 ({len(chunk)}件)')


# ==================================================================================================
# 6. 案件アラートファイル一括送付
# ==================================================================================================
def create_aa_bulk():

    # 作業フォルダ作成
    BASE_FOLDER_SEC = BASE_FOLDER / f'案件アラートファイル一括送付'
    WORKING_FOLDER = BASE_FOLDER_SEC / YYYYMMDD
    INPUT_FOLDER = WORKING_FOLDER / 'インプットデータ'

    if not WORKING_FOLDER.exists():
        make_folders([WORKING_FOLDER, INPUT_FOLDER])
        print("フォルダを作成しました. CSVを格納後、再度処理を実行してください")
        return


    # ファイルの取得
    send_df = selectfile_to_df("案件アラート送付先リスト.csv", INPUT_FOLDER)
    person_df = selectfile_to_df("案件管理担当者リスト.csv", INPUT_FOLDER)
    seiyaku_df = selectfile_to_df("成約報告担当者リスト.csv", INPUT_FOLDER)
    chakko_df = selectfile_to_df("着工証跡担当者リスト.csv", INPUT_FOLDER)
    target_df = selectfile_to_df("案件アラート送付対象クライアント一覧.xlsx", INPUT_FOLDER)



    # --- 1. 案件管理結合 ---
    merged_df1 = pd.merge(
        send_df[['振分先コード', '案件管理主体の振分先コード', '企業名', '振分先名']], 
        person_df[['振分先コード', '担当者部署', '担当者氏名', '担当者メールアドレス']], 
        left_on='案件管理主体の振分先コード', 
        right_on='振分先コード', 
        how='inner'
    )
    # 列整理
    merged_df1 = merged_df1.rename(columns={'振分先コード_x': '振分先コード'})
    merged_df1 = merged_df1.drop(columns=['案件管理主体の振分先コード', '振分先コード_y'])

    # 出力
    base_name = f"{YYYYMMDD}_送付先リストx案件管理担当者リスト"
    save_filepath = serial_filepath(WORKING_FOLDER, base_name,'.csv')
    merged_df1.to_csv(save_filepath, index=False, encoding="CP932")


    # --- 2. 成約報告結合 ---
    merged_df2 = pd.merge(
        send_df[['振分先コード', '成約報告主体の振分先コード', '企業名', '振分先名']], 
        seiyaku_df[['振分先コード', '担当者部署', '担当者氏名', '担当者メールアドレス']], 
        left_on='成約報告主体の振分先コード', 
        right_on='振分先コード', 
        how='inner'
    )
    # 列整理
    merged_df2 = merged_df2.rename(columns={'振分先コード_x': '振分先コード'})
    merged_df2 = merged_df2.drop(columns=['成約報告主体の振分先コード', '振分先コード_y'])

    # 出力
    base_name = f"{YYYYMMDD}_送付先リストx成約担当者リスト"
    save_filepath = serial_filepath(WORKING_FOLDER, base_name,'.csv')
    merged_df2.to_csv(save_filepath, index=False, encoding="CP932")


    # --- 3. 着工証跡結合 --- 
    merged_df3 = pd.merge(
        send_df[['振分先コード', '着工報告主体の振分先コード', '企業名', '振分先名']], 
        chakko_df[['振分先コード', '担当者部署', '担当者氏名', '担当者メールアドレス']], 
        left_on='着工報告主体の振分先コード', 
        right_on='振分先コード', 
        how='inner'
    )
    # 列整理
    merged_df3 = merged_df3.rename(columns={'振分先コード_x': '振分先コード'})
    merged_df3 = merged_df3.drop(columns=['着工報告主体の振分先コード', '振分先コード_y'])

    # 出力
    base_name = f"{YYYYMMDD}_送付先リストx着工担当者リスト"
    save_filepath = serial_filepath(WORKING_FOLDER, base_name,'.csv')
    merged_df3.to_csv(save_filepath, index=False, encoding="CP932")


    # --- 4. 縦結合 --- 
    combined_df = pd.concat([merged_df1, merged_df2, merged_df3], axis=0)
    # 列整理
    combined_df = combined_df.drop_duplicates()


    # --- 5. ファイル名結合 --- 
    merged_df4 = pd.merge(
        combined_df, 
        target_df[['振分先コード', 'ファイル名']], 
        left_on='振分先コード', 
        right_on='振分先コード', 
        how='inner'
    )

    # 出力
    base_name = f"{YYYYMMDD}_送付先リストx全担当者リスト_ファイル名"
    save_filepath = serial_filepath(WORKING_FOLDER, base_name,'.csv')
    merged_df4 = merged_df4.replace('\u200b', '', regex=True)
    merged_df4.to_csv(save_filepath, index=False, encoding="CP932")


    # --- 6. 整形 --- 
    def create_name_format(row):
        kigyo = row['企業名']
        furi = row['振分先名']
        busho = row['担当者部署']
        shimei = row['担当者氏名']
        mail = row['担当者メールアドレス']
        
        # 振分先名がある場合だけカッコを付ける
        furi_part = f" ({furi}分)" if furi != "" else ""
        
        # 全体を組み立てる
        return f"{kigyo}{furi_part} {busho} {shimei} <{mail}>"

    # 新しいデータフレームを作成
    output_df = pd.DataFrame()
    output_df['名前'] = merged_df4.apply(create_name_format, axis=1)
    output_df['ファイル名'] = merged_df4['ファイル名']


    # --- 5. 出力 (1000件ごとに分割) --- 
    chunk_size = 1000
    base_name_template = f"{YYYYMMDD}_案件アラート宛先取込用"

    # 全件数がchunk_size(1000)を超えるかどうかでループ
    for i, start_idx in enumerate(range(0, len(output_df), chunk_size)):
        chunk = output_df[start_idx : start_idx + chunk_size]
        
        # 複数ファイルになる場合はファイル名に連番(_1, _2...)を付与
        if len(output_df) > chunk_size:
            base_name = f"{base_name_template}({i+1})"
        else:
            base_name = base_name_template
            
        save_filepath = serial_filepath(WORKING_FOLDER, base_name, '.csv')
        chunk.to_csv(save_filepath, index=False, encoding="CP932")

        print(f' >> {str(save_filepath.name)} を作成 ({len(chunk)}件)')



if __name__ == '__main__':
        
    try:
        # # ==== 案件管理シートの一括送付 ====
        # print(f'[{datetime.datetime.now() :%Y-%m-%d %H:%M:%S}] 案件管理シートの一括送付リスト作成開始')
        # create_aks_bulk()
        # print(f'[{datetime.datetime.now() :%Y-%m-%d %H:%M:%S}] 案件管理シートの一括送付リスト作成終了')

        # # ==== 成約確認書・着工証跡引き取り便送付 ====
        # print(f'[{datetime.datetime.now() :%Y-%m-%d %H:%M:%S}] 成約確認書・着工証跡引き取り便送付リスト作成開始')
        # create_doc_pickup()
        # print(f'[{datetime.datetime.now() :%Y-%m-%d %H:%M:%S}] 成約確認書・着工証跡引き取り便送付リスト作成終了')


        # # ==== 個別着工証跡引き取り便送付 ====
        # print(f'[{datetime.datetime.now() :%Y-%m-%d %H:%M:%S}] 個別着工証跡引き取り便送付リスト作成開始')
        # create_doc_pickup_indv()
        # print(f'[{datetime.datetime.now() :%Y-%m-%d %H:%M:%S}] 個別着工証跡引き取り便送付リスト作成終了')


        # # ==== 案件管理シートリマインド ====
        # print(f'[{datetime.datetime.now() :%Y-%m-%d %H:%M:%S}] 案件管理シートリマインド送付リスト作成開始')
        # create_aks_remind()
        # print(f'[{datetime.datetime.now() :%Y-%m-%d %H:%M:%S}] 案件管理シートリマインド送付リスト作成終了')


        # # ==== HONEY管理クライアントへの案件進捗メール一括送付 ====
        # print(f'[{datetime.datetime.now() :%Y-%m-%d %H:%M:%S}] HONEY管理クライアントへの案件進捗メール一括送付リスト作成開始')
        # create_Honey_progress()
        # print(f'[{datetime.datetime.now() :%Y-%m-%d %H:%M:%S}] HONEY管理クライアントへの案件進捗メール一括送付リスト作成終了')


        # # ==== 案件アラートファイル一括送付 ====
        # print(f'[{datetime.datetime.now() :%Y-%m-%d %H:%M:%S}] 案件アラートファイル一括送付リスト作成開始')
        # create_aa_bulk()
        # print(f'[{datetime.datetime.now() :%Y-%m-%d %H:%M:%S}] 案件アラートファイル一括送付リスト作成終了')


        pass

    except KeyboardInterrupt:
        print('\n⛔ 処理を中断しました')    
    except Exception as e:
        err_name = type(e).__name__
        err_msg = str(e)
        print(f"\n⚠️ エラーが発生しました \nエラー名: {err_name}\n詳細: {err_msg}")
        import traceback
        traceback.print_exc()