# -*- coding: utf-8 -*-
"""
Project: サンレモ成約捕捉
File: utils.py
Description: 
    共通関数群および独自例外の定義
    1. システム共通の独自例外クラス（AppError系）の定義
    2. config.yml の自動読み込み
    3. 外部エクセル（PW一覧）からのパスワード取得
    4. ファイル名などの作成
    5. ダイアログによるファイル・フォルダ選択
    6. ファイル名にれんばん
    7. Windows禁止文字のサニタイズ（ファイル名浄化）

Copyright (c) 2026 SCSK ServiceWare Corporation.
All rights reserved.
"""
import sys
import re
import datetime
import shutil
from pathlib import Path
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import yaml
from itertools import count

YYYYMMDD = datetime.date.today().strftime("%Y%m%d")

# ============================================================================
# エラー定義
# ============================================================================
class AppError(Exception):
    pass
class ZeroDataError(AppError):
    """データ件数が0"""
    pass
class NotSelectError(AppError):
    """ファイル・フォルダを選択せずにダイアログ閉じた"""
    pass
class  ReferencePathError(AppError):
    """外部参照ファイルの中のパスが存在しない"""
    pass
class NoSheetError(AppError):
    """シートが存在しない"""
    pass

# ============================================================================
# 実行ファイルの場所取得（exe化対応）
# ============================================================================
def get_base_dir():
    # .exe化されているかチェック
    if getattr(sys, 'frozen', False):
        # 実行ファイル (.exe) の場所
        return Path(sys.executable).parent
    else:
        # スクリプト (.py) の場所
        return Path(__file__).parent

# ============================================================================
# 外部参照ファイル(config)の読み込み、ファイルパスが存在するかチェック
# ============================================================================
def set_config() -> dict:
    # exe化後も config.yml を実行ファイルと同じフォルダから読み込む
    config_path = get_base_dir() / "config.yml"
    with open(config_path, encoding="utf-8") as f:
        config = yaml.safe_load(f)

    not_exists = []

    # ルートフォルダが存在するか
    if 'ルートフォルダ' in config:
        root_folder = Path(config['ルートフォルダ'])
        if not root_folder.is_dir():
            not_exists.append(f"【ルートフォルダ】{config['ルートフォルダ']}")

    # 各ファイルが存在するか
    if 'ファイルパス' in config:
        for k, filepath in config['ファイルパス'].items():
            path = Path(filepath)
            if not path.exists():
                not_exists.append(f"【{k}】{filepath}")

    if not_exists:
        raise ReferencePathError(f"以下のファイルパスが見つかりません:\n" + "\n".join(not_exists))
    
    return config


# ============================================================================
# 作業フォルダの存在チェック
# ============================================================================
def exist_folders(folder_list: list):
    all_exist = all([folder.exists() for folder in folder_list])

    if not all_exist:
        raise ReferencePathError(f'作業フォルダーが作成されていません。フォルダーを作成してください')
    

# ============================================================================
# パスワード取得
# ============================================================================
def get_pw(code:str) -> str:
    config = set_config()
    pw_filepath = config['ファイルパス']['PW一覧']
    pw_df = pd.read_excel(pw_filepath, dtype=str)
    pw_series = pw_df.loc[pw_df['振分先コード'] == code, 'PW']

    if pw_series.empty or len(pw_series) > 1:
        return None
    
    pw = pw_series.iloc[0]
    return pw 

# ============================================================================
# 企業名、ファイル名生成
# ============================================================================
def create_name(kigyo_nm, furiwakesaki_nm=None):
    if furiwakesaki_nm:
        return f"{kigyo_nm}様（{furiwakesaki_nm}分）"
    else:
        return f"{kigyo_nm}様"

def create_filename(code, kigyo_nm, furiwakesaki_nm=None):
    term = datetime.date.today().strftime("%Y年%m月")
    name = create_name(kigyo_nm, furiwakesaki_nm)
    return f"【SUUMO注文】{term}度案件管理シート_{code}_{name}.xlsx"

def create_filename_alert(code, kigyo_nm, furiwakesaki_nm=None):
    name = create_name(kigyo_nm, furiwakesaki_nm)
    return f"【SUUMO注文】{YYYYMMDD}_(確認用)案件アラート_{code}_{name}.xlsx"


# ============================================================================
# フォルダ・ファイルの取得
# ============================================================================
# --- 1. ダイアログでフォルダを取得 ---
def select_folder(title:str="フォルダを選択してください", initial_dir:Path=None) -> Path:
    root = tk.Tk()
    root.withdraw()
    # 1. 最前面に設定
    root.attributes('-topmost', True)
    # 2. 強制的にフォーカスを当てる
    root.focus_force()
    folder_path = filedialog.askdirectory(title=title, initialdir=initial_dir)
    root.destroy()
    if not folder_path: 
        raise NotSelectError('フォルダが選択されませんでした')

    return Path(folder_path)

# --- 2. ダイアログでファイルを取得 ---
def select_file(title:str='ファイルを選択してください', file_types:list=[("すべてのファイル", "*.*")], initial_dir:Path=None) -> Path:
    root = tk.Tk()
    root.withdraw()
    # 1. 最前面に設定
    root.attributes('-topmost', True)
    # 2. 強制的にフォーカスを当てる
    root.focus_force()
    file_path = filedialog.askopenfilename(title=title, filetypes=file_types, initialdir=initial_dir)
    root.destroy()
    if not file_path: 
        raise NotSelectError('ファイルが選択されませんでした')

    return Path(file_path)

# --- 3. 指定フォルダから指定ワードでファイルを取得(複数あった場合はダイアログで) ---
def search_file(folder_path:Path, file_name:str) -> Path:
    files = [f for f in folder_path.glob(f"*{file_name}*") if not f.name.startswith('~$')]
    if len(files) == 1:
        return files[0]
    else:
        print(f'{file_name}を選択してください')
        return select_file(initial_dir=folder_path)


# --- 4. 指定フォルダから指定のファイルを取得しdfで返す ---
def selectfile_to_df(filename:str, initial_dir:Path=None) -> pd.DataFrame:
    """dfはfillna('') | filenameに拡張子必要"""
    
    print(f'{filename}を選択してください')
    
    # 拡張子による条件分岐
    if filename.endswith('.csv'):
        filepath = select_file(title=f'{filename}を選択してください', file_types=[('CSVファイル', '*.csv')], initial_dir=initial_dir)
        df = pd.read_csv(filepath, encoding='CP932', dtype=str).fillna('')
    
    elif filename.endswith('.xlsx'):
        filepath = select_file(title=f'{filename}を選択してください', file_types=[('EXCELファイル', '*.xlsx')], initial_dir=initial_dir)
        df = pd.read_excel(filepath, dtype=str).fillna('')

    else:
        print("対応していないファイル形式です")
        return None

    df = df.map(lambda x: x.strip() if isinstance(x, str) else x) # データクレンジング
    
    print(f' << {filepath} を取得')
    return df


# ============================================================================
# その他の共通関数
# ============================================================================
# --- ファイル名に連番付与 ---
def serial_filepath(folder:Path, base_name:str, ext:str):
    if not ext.startswith("."): ext = '.' + ext
    return next(
        path for i in count()
        if not (path := folder / (f"{base_name}_{i}{ext}" if i > 0 else f"{base_name}{ext}")).exists()
    )

# --- windowsファイル名禁止文字削除 ---
def sanitize_filename(name: str, replacement: str = "") -> str:
    # 禁止文字を置換
    name = re.sub(r'[\\/:*?"<>|]', replacement, name)

    # 末尾のドット・スペースを削除
    name = name.rstrip(". ")

    # 空になった場合の保険
    if not name:
        name = "禁止文字のみのパス"
    
    return name