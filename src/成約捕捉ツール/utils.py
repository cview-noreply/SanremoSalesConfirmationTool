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
    6. ファイル名に連番
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


# エラー無視
import warnings
warnings.simplefilter('ignore', UserWarning)


YYYYMMDD = datetime.date.today().strftime('%Y%m%d')

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
class PICError(AppError):
    """担当者列に1が立っていない"""
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
    config_path = get_base_dir() / 'config.yml'
    with open(config_path, encoding='utf-8') as f:
        config = yaml.safe_load(f)

    not_exists = []

    # ルートフォルダが存在するか
    if 'ルートフォルダ' in config:
        root_folder = Path(config['ルートフォルダ'])
        if not root_folder.is_dir():
            not_exists.append(f'【ルートフォルダ】{config['ルートフォルダ']}')

    # 各ファイルが存在するか
    if 'ファイルパス' in config:
        for k, filepath in config['ファイルパス'].items():
            path = Path(filepath)
            if not path.exists():
                not_exists.append(f'【{k}】{filepath}')

    if not_exists:
        raise ReferencePathError(f'以下のファイルパスが見つかりません:\n' + '\n'.join(not_exists))
    
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
def create_name(kigyo_nm, furi_nm=None):
    if furi_nm:
        return f'{kigyo_nm}様({furi_nm}分)' # 半角かっこ
    else:
        return f'{kigyo_nm}様'

def create_filename_anken(code, kigyo_nm, furi_nm=None):
    """【SUUMO注文】2026年04月度_案件管理シート_01234567890_かもめ工務店様(A支店分)"""
    term = datetime.date.today().strftime('%Y年%m月')
    name = create_name(kigyo_nm, furi_nm)
    return f'【SUUMO注文】{term}度_案件管理シート_{code}_{name}.xlsx'

def create_filename_alert(code, kigyo_nm, furi_nm=None):
    """【SUUMO注文】20260402_(確認用)案件アラート_01234567890_かもめ工務店様(A支店分)"""
    name = create_name(kigyo_nm, furi_nm)
    return f'【SUUMO注文】{YYYYMMDD}_(確認用)案件アラート_{code}_{name}.xlsx'

# 送付用リストのファイル名の作成
def create_filename(kigyo_nm:str=None, furi_nm:str=None, busho:str=None, shimei:str=None, mailaddress:str=None):
    # 企業名(振分先名分)
    if furi_nm:
        kigyo_part = f'{kigyo_nm}様　({furi_nm}分)'
    else:
        kigyo_part = f'{kigyo_nm}様'

    # 氏名<メールアドレス>
    if mailaddress:
        person_part = f'{shimei}<{mailaddress}>'
    else:
        person_part = f'{shimei}'

    # 部署の有無
    if busho:
        filename = f'{kigyo_part}　{busho}　{person_part}'
    else:
        filename = f'{kigyo_part}　{person_part}'
    
    return filename
    
# ============================================================================
# フォルダ・ファイルの取得
# ============================================================================
# --- 1. ダイアログでフォルダを取得 ---
def select_folder(title:str='フォルダを選択してください', initial_dir:Path=None) -> Path:
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
def select_file(title:str='ファイルを選択してください', file_types:list=[('すべてのファイル', '*.*')], initial_dir:Path=None) -> Path:
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
    files = [f for f in folder_path.glob(f'*{file_name}*') if not f.name.startswith('~$')]
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
        print('対応していないファイル形式です')
        return None

    df = df.map(lambda x: x.strip() if isinstance(x, str) else x) # データクレンジング
    
    print(f' << {filepath} を取得')
    return df


# ============================================================================
# その他の共通関数
# ============================================================================
# --- ファイル名に連番付与 ---
def serial_filepath(folder:Path, base_name:str, ext:str):
    if not ext.startswith('.'): ext = '.' + ext
    return next(
        path for i in count()
        if not (path := folder / (f'{base_name}_{i}{ext}' if i > 0 else f'{base_name}{ext}')).exists()
    )

# --- windowsファイル名禁止文字削除 ---
def sanitize_filename(name: str, replacement: str = '') -> str:
    # 禁止文字を置換
    name = re.sub(r"[\\/:*?'<>|]", replacement, name)

    # 末尾のドット・スペースを削除
    name = name.rstrip('. ')

    # 空になった場合の保険
    if not name:
        name = '禁止文字のみのパス'
    
    return name


# ============================================================================
# 送付用リストのチェック
# ============================================================================
class ListChecker:
    def __init__(self, created_df: pd.DataFrame, source_df: pd.DataFrame):
        '''
        created_df: 取込用リスト（DataFrameとして受け取る）
        source_df:  送付先リストx担当者リスト
        '''
        self.created_df = created_df  # ← ファイルパスではなくDF
        self.source_df = source_df

    def check_has_file(self, save_filepath):
        '''
        名前: 企業名　(振分先名分)　部署名　担当者氏名<担当者メールアドレス> ※全角スペース、半角かっこ
        ファイル名: 【SUUMO注文】YYYY年M月度_XXXXXX_{振分先コード}_{企業名}様({振分先名}分).xlsx
        '''
        result = []

        df = self.created_df.copy()
        df.columns = ['名前', 'ファイル名']  # header=Noneで保存したCSVに合わせて列名を付与

        for row in df.itertuples():
            report = {}
            report['名前'] = row.名前
            report['ファイル名'] = row.ファイル名

            try:
                # ────── 名前の分割 ──────
                n_spl = row.名前.split('　')
                n_corpname = n_spl[0].rstrip('様')

                if len(n_spl) == 2: # 振分先名なし、部署なし
                    n_furiname = ''
                    n_busho = ''
                elif len(n_spl) == 3: # 振分先名なし、もしくは部署なし
                    if '分)' in n_spl[1]:
                        n_furiname = n_spl[1].replace('(', '').replace('分)','')
                        n_busho = ''
                    else:
                        n_furiname = ''
                        n_busho = n_spl[1]
                elif len(n_spl) == 4: # 振分先名あり、部署あり
                    n_furiname = n_spl[1].replace('(', '').replace('分)','')
                    n_busho = n_spl[2]

                n_prsnname = n_spl[-1].split('<')[0]
                n_mailaddress = n_spl[-1].split('<')[1][:-1] # >を除く

                # n_corpname, n_furiname, n_busho, n_prsnname, n_mailaddress

                #  ────── ファイル名の分割 ──────
                f_spl = row.ファイル名.split('_')
                furicode = f_spl[2]
                f_kigyo_part = f_spl[3].replace('.xlsx', '').split('(')
                f_corpname = f_kigyo_part[0].rstrip('様')

                if len(f_kigyo_part) == 2:
                    f_furiname = f_kigyo_part[1].replace('分)', '')
                else:
                    f_furiname = ''

                # furicode, f_corpname, f_furiname

                # ────── チェック ──────
                if n_corpname != f_corpname or n_furiname != f_furiname:
                    report['チェック結果'] = '名前とファイル名の企業名/振分先名相違'
                
                matched_rows = self.source_df[
                    (self.source_df['振分先コード'].fillna('') == furicode) & # 送付先リストの振分先コード
                    (self.source_df['企業名'].fillna('') == f_corpname) &
                    (self.source_df['振分先名'].fillna('') == f_furiname) & 
                    (self.source_df['担当者部署'].fillna('') == n_busho) &
                    (self.source_df['担当者氏名'].fillna('') == n_prsnname) &
                    (self.source_df['担当者メールアドレス'].fillna('') == n_mailaddress)
                ]
                if matched_rows.empty:
                    report['チェック結果'] = '不備あり'
                    report['振分先コード'] = furicode
                    report['企業名'] = f_corpname
                    report['振分先名'] = f_furiname
                    report['担当者部署'] = n_busho
                    report['担当者氏名'] = n_prsnname
                    report['担当者メールアドレス'] = n_mailaddress
                else:
                    report['チェック結果'] = '不備なし'


                    
            except Exception as e:
                report['チェック結果'] = 'チェック時エラー'
                print(e)
            
            finally:
                result.append(report)
            
        result_df = pd.DataFrame(result)
        result_df.index = result_df.index + 1
        result_df.to_excel(save_filepath)

    def check_no_file(self, save_filepath):
        '''
        名前: 企業名　(振分先名)　部署名　担当者氏名
        メールアドレス: メールアドレス
        '''
        result = []

        df = self.created_df.copy()

        for row in df.itertuples():
            report = {}
            report['名前'] = row.名前
            report['メールアドレス'] = row.メールアドレス

            try:
                # ────── 名前の分割 ──────
                n_spl = row.名前.split('　')
                n_corpname = n_spl[0].rstrip('様')

                if len(n_spl) == 2: # 振分先名なし、部署なし
                    n_furiname = ''
                    n_busho = ''
                elif len(n_spl) == 3: # 振分先名なし、もしくは部署なし
                    if '分)' in n_spl[1]:
                        n_furiname = n_spl[1].replace('(', '').replace('分)','')
                        n_busho = ''
                    else:
                        n_furiname = ''
                        n_busho = n_spl[1]
                elif len(n_spl) == 4: # 振分先名あり、部署あり
                    n_furiname = n_spl[1].replace('(', '').replace('分)','')
                    n_busho = n_spl[2]

                n_prsnname = n_spl[-1]

                # n_corpname, n_furiname, n_busho, n_prsnname

                #  ────── メールアドレス ──────
                mailaddress = row.メールアドレス

                # ────── チェック ──────
                matched_rows = self.source_df[
                    (self.source_df['企業名'].fillna('') == n_corpname) &
                    (self.source_df['振分先名'].fillna('') == n_furiname) & 
                    (self.source_df['担当者部署'].fillna('') == n_busho) &
                    (self.source_df['担当者氏名'].fillna('') == n_prsnname) &
                    (self.source_df['担当者メールアドレス'].fillna('') == mailaddress)
                ]
                if matched_rows.empty:
                    report['チェック結果'] = '不備あり'
                    report['企業名'] = n_corpname
                    report['振分先名'] = n_furiname
                    report['担当者部署'] = n_busho
                    report['担当者氏名'] = n_prsnname
                    report['担当者メールアドレス'] = mailaddress
                else:
                    report['チェック結果'] = '不備なし'


                    
            except Exception as e:
                report['チェック結果'] = 'チェック時エラー'
            
            finally:
                result.append(report)
            
        result_df = pd.DataFrame(result)
        result_df.index = result_df.index + 1
        result_df.to_excel(save_filepath)


    def check_kobetsu_file(self, save_filepath):
        '''
        名前: 企業名　(振分先名分)　部署名　担当者氏名<担当者メールアドレス> ※全角スペース、半角かっこ
        ファイル名: 【SUUMO注文】YYYY年M月度_XXXXXX_{振分先コード}_{企業名}様({振分先名}分).xlsx
        '''
        result = []

        df = self.created_df.copy()
        df.columns = ['名前', 'ファイル名']  # header=Noneで保存したCSVに合わせて列名を付与

        for row in df.itertuples():
            report = {}
            report['名前'] = row.名前
            report['ファイル名'] = row.ファイル名

            try:
                # ────── 名前の分割 ──────
                n_spl = row.名前.split('　')
                n_corpname = n_spl[0].rstrip('様')

                if len(n_spl) == 2: # 振分先名なし、部署なし
                    n_furiname = ''
                    n_busho = ''
                elif len(n_spl) == 3: # 振分先名なし、もしくは部署なし
                    if '分)' in n_spl[1]:
                        n_furiname = n_spl[1].replace('(', '').replace('分)','')
                        n_busho = ''
                    else:
                        n_furiname = ''
                        n_busho = n_spl[1]
                elif len(n_spl) == 4: # 振分先名あり、部署あり
                    n_furiname = n_spl[1].replace('(', '').replace('分)','')
                    n_busho = n_spl[2]

                n_prsnname = n_spl[-1].split('<')[0]
                n_mailaddress = n_spl[-1].split('<')[1][:-1] # >を除く

                # n_corpname, n_furiname, n_busho, n_prsnname, n_mailaddress

                #  ────── ファイル名の分割 ──────
                f_spl = row.ファイル名.split('_')
                furicode = f_spl[2]
                f_kigyo_part = f_spl[3].replace('.xlsx', '').split('(')
                f_corpname = f_kigyo_part[0].rstrip('様')

                if len(f_kigyo_part) == 2:
                    f_furiname = f_kigyo_part[1].replace('分)', '')
                else:
                    f_furiname = ''

                # furicode, f_corpname, f_furiname

                # ────── チェック ──────
                if n_corpname != f_corpname or n_furiname != f_furiname:
                    report['チェック結果'] = '名前とファイル名の企業名/振分先名相違'
                
                matched_rows = self.source_df[
                    (self.source_df['振分先コード'].fillna('') == furicode) & # 送付先リストの振分先コード
                    (self.source_df['企業名'].fillna('') == f_corpname) &
                    (self.source_df['振分先名'].fillna('') == f_furiname) & 
                    (self.source_df['この案件のみの着工報告担当者部署'].fillna('') == n_busho) &
                    (self.source_df['この案件のみの着工報告担当者氏名'].fillna('') == n_prsnname) &
                    (self.source_df['この案件のみの着工報告担当者メールアドレス'].fillna('') == n_mailaddress)
                ]
                if matched_rows.empty:
                    report['チェック結果'] = '不備あり'
                    report['振分先コード'] = furicode
                    report['企業名'] = f_corpname
                    report['振分先名'] = f_furiname
                    report['担当者部署'] = n_busho
                    report['担当者氏名'] = n_prsnname
                    report['担当者メールアドレス'] = n_mailaddress
                else:
                    report['チェック結果'] = '不備なし'


                    
            except Exception as e:
                report['チェック結果'] = 'チェック時エラー'
                print(e)
            
            finally:
                result.append(report)
            
        result_df = pd.DataFrame(result)
        result_df.index = result_df.index + 1
        result_df.to_excel(save_filepath)