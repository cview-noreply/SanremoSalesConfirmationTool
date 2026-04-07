# -*- coding: utf-8 -*-
"""
Project: サンレモ成約捕捉
File: app.py
Description: 
    1. 各処理用ボタンを配置したGUIの生成 (tkinter使用)

Copyright (c) 2026 SCSK ServiceWare Corporation.
All rights reserved.
"""

import os
import sys
import threading
import datetime
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
from pathlib import Path

# 各自作モジュールのインポート
import create_sheets
import check_sheets
import create_sendlist
import store_documents
import create_alert

from utils import get_base_dir

class LogRedirect:
    def __init__(self, text_widget, log_list):
        self.text_widget = text_widget
        self.log_list = log_list

    def write(self, text):
        msg = text.rstrip()
        if msg:
            now = datetime.datetime.now().strftime('%H:%M:%S')
            full_msg = f'[{now}] {msg}'
            
            # GUIスレッドでテキストエリアを更新
            self.text_widget.after(0, self._update_ui, full_msg)
            self.log_list.append(full_msg)

    def _update_ui(self, msg):
        self.text_widget.configure(state='normal')
        self.text_widget.insert(tk.END, msg + '\n')
        self.text_widget.see(tk.END)
        self.text_widget.configure(state='disabled')

    def flush(self):
        pass

class SanremoApp:
    def __init__(self, root):
        self.root = root
        self.root.title('サンレモ成約捕捉 処理ツール')
        self.root.geometry('1200x700')
        self._log_lines = []

        # アイコン設定 
        try:
            # exe化後はアイコンが同じフォルダに展開される
            icon_path = get_base_dir() / 'icon.ico'
            self.root.iconbitmap(str(icon_path))
        except Exception:
            pass  # アイコンがなくても落ちないようにする
        
        # スタイル設定
        style = ttk.Style()
        style.configure('TButton', font=('Meiryo', 10), padding=5)
        style.configure('Primary.TButton', font=('Meiryo', 10, 'bold'), foreground='#1a6fb5')
        style.configure('Title.TLabel', font=('Meiryo', 18, 'bold'), foreground='#1a6fb5')

        self.setup_ui()
        
        # 標準出力のリダイレクト
        self.redirector = LogRedirect(self.log_area, self._log_lines)
        sys.stdout = self.redirector
        sys.stderr = self.redirector

    def setup_ui(self):
        # メインコンテナ
        main_frame = ttk.Frame(self.root, padding='20')
        main_frame.pack(fill=tk.BOTH, expand=True)

        # タイトル
        title_label = ttk.Label(main_frame, text='サンレモ成約捕捉 処理ツール', style='Title.TLabel')
        title_label.pack(pady=(0, 20))

        # 左右分割用パンウィンドウ
        paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # --- 左カラム: タブエリア ---
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=1)

        self.notebook = ttk.Notebook(left_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.create_tab1_sheets()
        self.create_tab2_check()
        self.create_tab3_list()
        self.create_tab4_store()
        self.create_tab5_alert()

        # --- 右カラム: ログエリア ---
        right_frame = ttk.Frame(paned, padding=(20, 0, 0, 0))
        paned.add(right_frame, weight=1)

        log_header = ttk.Frame(right_frame)
        log_header.pack(fill=tk.X)
        
        ttk.Label(log_header, text='実行ログ:', font=('Meiryo', 10, 'bold')).pack(side=tk.LEFT)
        
        save_btn = ttk.Button(log_header, text='💾 ログをテキストで保存', command=self.save_log_to_file)
        save_btn.pack(side=tk.RIGHT)

        self.log_area = scrolledtext.ScrolledText(
            right_frame, bg='#cccccc', fg='#222222', 
            font=('Consolas', 10), state='disabled', wrap=tk.WORD
        )
        self.log_area.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

    def create_tab_content(self, title):
        frame = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(frame, text=title)
        
        lbl = ttk.Label(frame, text=title, font=('Meiryo', 12, 'bold'))
        lbl.pack(pady=(0, 15))
        
        container = ttk.Frame(frame)
        container.pack(fill=tk.BOTH, expand=True)
        return container

    def make_btn(self, parent, label, func, primary=False, task_name=''):
        name = task_name or label
        style = 'Primary.TButton' if primary else 'TButton'
        
        btn = ttk.Button(parent, text=label, style=style, 
                         command=lambda: self.start_thread(func, name))
        btn.pack(fill=tk.X, pady=4, padx=50)

    def start_thread(self, func, task_name):
        def wrapper():
            now = datetime.datetime.now().strftime('%H:%M:%S')
            print(f'▶ {task_name} 開始')
            try:
                func()
                print(f'✅ {task_name} 完了')
            except Exception as e:
                print(f' ⚠️ エラー発生: {e}')
        
        threading.Thread(target=wrapper, daemon=True).start()

    def save_log_to_file(self):
        try:
            file_path = get_base_dir() / f'実行ログ_{datetime.datetime.now().strftime('%Y%m%d%H%M')}.txt'
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(self._log_lines))
            messagebox.showinfo('成功', f'ログを保存しました:\n{file_path}')
        except Exception as e:
            messagebox.showerror('エラー', f'保存に失敗しました: {e}')

    # --- 各タブのレイアウト定義 ---
    def create_tab1_sheets(self):
        c = self.create_tab_content('1. 案件管理シート作成')
        self.make_btn(c, 'フォルダ作成', create_sheets.make_folders, task_name='案件管理シート作成 > フォルダ作成')
        self.make_btn(c, '案件管理シート生成', create_sheets.create_sheets, task_name='案件管理シート作成 > 案件管理シート生成')
        self.make_btn(c, '作成後チェック', create_sheets.check_sheets_at_creation, task_name='案件管理シート作成 > 作成後チェック')
        self.make_btn(c, '送付対象CL一覧作成', create_sheets.create_send_list, task_name='案件管理シート作成 > 送付対象CL一覧作成')
        self.make_btn(c, '送信実績取込用ファイル作成', create_sheets.create_send_jisseki, task_name='案件管理シート作成 > 送信実績取込用ファイル作成')

    def create_tab2_check(self):
        c = self.create_tab_content('2. 案件管理シート受理チェック')
        self.make_btn(c, 'フォルダ作成', check_sheets.make_folders, task_name='受理チェック > フォルダ作成')
        self.make_btn(c, '不備チェック実行', check_sheets.check_sheets_on_receipt, task_name='受理チェック > 不備チェック実行')
        self.make_btn(c, '回収実績取込用ファイル作成', check_sheets.create_receive_jisseki, task_name='受理チェック > 回収実績取込用ファイル作成')
        self.make_btn(c, '受領メール宛先作成', check_sheets.create_mail_list, task_name='受理チェック > 受領メール宛先作成')
        self.make_btn(c, '反響情報CSV作成', check_sheets.create_input_data_hankyo, task_name='受理チェック > 反響情報CSV作成')

    def create_tab3_list(self):
        c = self.create_tab_content('3. 送付用リスト作成')
        self.make_btn(c, '案件管理シート一括送付用', create_sendlist.create_aks_bulk, task_name='送付リスト作成 > 案件管理シート一括送付用')
        self.make_btn(c, '成約/着工引き取り便送付用', create_sendlist.create_doc_pickup, task_name='送付リスト作成 > 成約/着工引き取り便送付用')
        self.make_btn(c, '個別着工引き取り便送付用', create_sendlist.create_doc_pickup_indv, task_name='送付リスト作成 > 個別着工引き取り便送付用')
        self.make_btn(c, '案件管理シート未提出PUSH送付用', create_sendlist.create_aks_remind, task_name='送付リスト作成 > 案件管理シート未提出PUSH送付用')
        self.make_btn(c, 'HONEY案件進捗送付用', create_sendlist.create_honey_progress, task_name='送付リスト作成 > HONEY案件進捗送付用')
        self.make_btn(c, '案件アラート一括送付用', create_sendlist.create_aa_bulk, task_name='送付リスト作成 > 案件アラート一括送付用')

    def create_tab4_store(self):
        c = self.create_tab_content("4. 証跡振り分け")

        def run_store():
            store_documents.make_folders()
            store_documents.store_documents()

        self.make_btn(c, "振り分け実行", run_store, task_name="証跡振り分け 実行")
        
    def create_tab5_alert(self):
        c = self.create_tab_content('5. 案件アラート作成')
        self.make_btn(c, 'フォルダ作成', create_alert.make_folders, task_name='案件アラート作成 > フォルダ作成')
        self.make_btn(c, '案件アラート作成', create_alert.create_sheets, task_name='案件アラート作成 > 案件アラート作成')
        self.make_btn(c, '作成後チェック', create_alert.check_sheets_at_creation, task_name='案件アラート作成 > 作成後チェック')
        self.make_btn(c, '送付対象CL一覧作成', create_alert.make_send_list, task_name='案件アラート作成 > 送付対象CL一覧作成')

if __name__ == '__main__':
    root = tk.Tk()
    app_gui = SanremoApp(root)
    root.mainloop()