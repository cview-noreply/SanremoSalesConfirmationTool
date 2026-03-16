# -*- coding: utf-8 -*-
"""
Project: サンレモ成約捕捉
File: app.py
Description: 
    1. 各処理用ボタンを配置したGUIの生成
    ※nice guiでUI作成

Copyright (c) 2026 SCSK ServiceWare Corporation.
All rights reserved.
"""
import os
import sys
import threading
import datetime
from nicegui import app, ui

# 各自作モジュールのインポート
import create_sheets
import check_sheets
import create_sendlist
import store_documents
import create_alert

from utils import get_base_dir

class LogRedirect:
    def __init__(self, log_element):
        self.log = log_element

    def write(self, text):
        # printの末尾にある改行を取り除き、中身がある場合のみ処理
        msg = text.rstrip()
        if msg:
            # 1. 画面のログエリアに表示
            self.log.push(msg)
            # 2. 保存用リストにタイムスタンプ付きで記録
            _log_lines.append(f"[{datetime.datetime.now():%H:%M:%S}] {msg}")

    def flush(self):
        pass


# ログの全テキストを保持するリスト（コピー用）
_log_lines: list[str] = []


def start_thread(func, log, task_name: str = ""):
    def wrapper():
        try:
            msg = f"[{datetime.datetime.now():%H:%M:%S}] ▶ {task_name} 開始"
            log.push(msg)
            _log_lines.append(msg)
            func()
            msg = f"[{datetime.datetime.now():%H:%M:%S}] ✅ {task_name} 完了"
            log.push(msg)
            _log_lines.append(msg)
        except Exception as e:
            msg = f"[{datetime.datetime.now():%H:%M:%S}] ⚠️ エラー発生: {e}"
            log.push(msg)
            _log_lines.append(msg)
    threading.Thread(target=wrapper, daemon=True).start()


def make_button(label, func, log, primary=False, task_name: str = ""):
    """task_name を省略すると label をそのまま使用"""
    name = task_name or label
    style = "width: 380px;"
    if primary:
        ui.button(label, on_click=lambda: start_thread(func, log, name)) \
            .style(style + 'margin:8px auto;')
    else:
        ui.button(label, on_click=lambda: start_thread(func, log, name)) \
            .props("outline") \
            .style(style + 'margin:0px auto;')


@ui.page("/")
def main():
    ui.query("body").style("background-color: #f0f0f0;")

    # ── タイトル ──────────────────────────────────────
    with ui.element("div").style("text-align:center; padding: 16px;"):
        ui.label("サンレモ成約捕捉 処理ツール") \
            .style("font-size: 24px; font-weight: bold; color: #1a6fb5;")

    # ── レイアウト用コンテナの準備 ──────────────────────
    with ui.row().style("width: 100%; max-width: 1440px; margin: 0 auto; justify-content: center; gap: 40px;"):
        left_column = ui.column().style("width: 50%; min-width: 400px;")
        right_column = ui.column().style("width: 46%; min-width: 400px;")

    # ── ログエリア（右カラムに配置） ──────────────────────
    with right_column:
        # ラベル＋コピーボタンを横並びに
        with ui.row().style("align-items: center; margin-bottom: 4px; gap: 12px;"):
            ui.label("実行ログ:").style("font-weight: bold;")
            
            # ログを同じフォルダにテキストファイルとして直接保存する処理
            def save_log_to_file():
                try:
                    import os
                    # ツールと同じフォルダに "execution_log.txt" を作成
                    file_path = get_base_dir() / f"実行ログ_{datetime.datetime.now().strftime('%Y%m%d%H%M')}.txt"
                    with open(file_path, "w", encoding="utf-8") as f:
                        f.write("\n".join(_log_lines))
                    
                    # 画面上に成功メッセージをポップアップ表示
                    ui.notify(f"ログを保存しました: {file_path}", type='positive', position='top')
                except Exception as e:
                    ui.notify(f"保存に失敗しました: {e}", type='negative', position='top')

            # テキスト保存ボタンを配置
            ui.button("💾 ログをテキストで保存", on_click=save_log_to_file) \
                .props("flat dense size=sm") \
                .style("font-size: 12px;")

        log = ui.log(max_lines=200) \
            .style("width: 100%; height: 400px; background: #cccccc; color: #222222;"
                   " font-family: Consolas, monospace; font-size: 13px;"
                   " border-radius: 6px; padding: 8px; overflow: auto; white-space: pre;") \
            .props('id="log-area"')

        # 標準出力をリダイレクト
        sys.stdout = LogRedirect(log)
        sys.stderr = LogRedirect(log)

        def cleanup():
            sys.stdout = sys.__stdout__
            sys.stderr = sys.__stderr__
        app.on_shutdown(cleanup)

    # ── タブエリア（左カラムに配置） ────────────────────
    with left_column:
        with ui.tabs().style("width: 100%;") as tabs:
            t1 = ui.tab("案件管理シート作成")
            t2 = ui.tab("受理チェック")
            t3 = ui.tab("送付リスト作成")
            t4 = ui.tab("証跡振り分け")
            t5 = ui.tab("案件アラート")

        with ui.tab_panels(tabs, value=t1).style(
            "width: 100%; min-height: 360px; background-color: #ffffff;"
            " display: flex; justify-content: center;"
        ):
            # ── タブ1: シート作成 ─────────────────────────
            with ui.tab_panel(t1):
                ui.label("1. 案件管理シート作成").style(
                    "font-size: 18px; font-weight: bold; margin-bottom: 8px; width: 380px; margin: 0 auto;")
                make_button("【一括】全工程を実行",
                            lambda: (
                                create_sheets.make_folders(),
                                create_sheets.create_sheets(),
                                create_sheets.check_sheets_at_creation(),
                                create_sheets.make_send_list(),
                                create_sheets.create_send_jisseki()
                            ),
                            log, primary=True, task_name="案件管理シート作成 全工程")
                make_button("フォルダ作成",               create_sheets.make_folders,              log, task_name="案件管理シート作成 > フォルダ作成")
                make_button("シート生成",                 create_sheets.create_sheets,             log, task_name="案件管理シート作成 > シート生成")
                make_button("整合性チェック",             create_sheets.check_sheets_at_creation,  log, task_name="案件管理シート作成 > 整合性チェック")
                make_button("送付対象リスト出力",         create_sheets.make_send_list,            log, task_name="案件管理シート作成 > 送付対象リスト出力")
                make_button("送信実績取込用ファイル作成", create_sheets.create_send_jisseki,       log, task_name="案件管理シート作成 > 送信実績取込用ファイル作成")

            # ── タブ2: 受理チェック ───────────────────────
            with ui.tab_panel(t2):
                ui.label("2. 受理チェック").style(
                    "font-size: 18px; font-weight: bold; margin-bottom: 8px; width: 380px; margin: 0 auto;")
                make_button("【一括】全工程を実行",
                            lambda: (
                                check_sheets.make_folders(),
                                check_sheets.check_sheets_on_receipt(),
                                check_sheets.make_input_data_cl(),
                                check_sheets.make_mail_list(),
                                check_sheets.make_input_data_hankyo()
                            ),
                            log, primary=True, task_name="受理チェック 全工程")
                make_button("フォルダ作成",   check_sheets.make_folders, log, task_name="受理チェック > フォルダ作成")
                make_button("不備チェック実行",   check_sheets.check_sheets_on_receipt, log, task_name="受理チェック > 不備チェック実行")
                make_button("回収実績CSV作成",   check_sheets.make_input_data_cl,      log, task_name="受理チェック > 回収実績CSV作成")
                make_button("受領メール宛先作成", check_sheets.make_mail_list,          log, task_name="受理チェック > 受領メール宛先作成")
                make_button("反響情報CSV作成",   check_sheets.make_input_data_hankyo,  log, task_name="受理チェック > 反響情報CSV作成")

            # ── タブ3: 送付リスト ─────────────────────────
            with ui.tab_panel(t3):
                ui.label("3. 送付リスト作成").style(
                    "font-size: 18px; font-weight: bold; margin-bottom: 8px; width: 380px; margin: 0 auto;")
                make_button("案件管理シート一括送付用",     create_sendlist.create_AKS_send,      log, task_name="送付リスト作成 > 案件管理シート一括送付用")
                make_button("成約/着工引き取り便作成",      create_sendlist.create_Doc_pickup,    log, task_name="送付リスト作成 > 成約/着工引き取り便")
                make_button("個別着工証跡作成",             create_sendlist.create_Doc_Indivisual,log, task_name="送付リスト作成 > 個別着工証跡")
                make_button("未提出リマインド作成",         create_sendlist.create_AKS_remind,    log, task_name="送付リスト作成 > 未提出リマインド")
                make_button("HONEY進捗依頼作成",            create_sendlist.create_Honey_progress,log, task_name="送付リスト作成 > HONEY進捗依頼")
                make_button("案件アラート一括送付用作成",   create_sendlist.create_AAF_send,      log, task_name="送付リスト作成 > 案件アラート一括送付用")

            # ── タブ4: 書類振分 ───────────────────────────
            with ui.tab_panel(t4):
                ui.label("4. 証跡振り分け").style(
                    "font-size: 18px; font-weight: bold; margin-bottom: 8px; width: 380px; margin: 0 auto;")
                make_button("振り分け実行",
                            lambda: (store_documents.make_folders(), store_documents.store_documents()),
                            log, primary=True, task_name="証跡振り分け 実行")
                make_button("フォルダ作成のみ", store_documents.make_folders, log, task_name="証跡振り分け > フォルダ作成")

            # ── タブ5: アラート ───────────────────────────
            with ui.tab_panel(t5):
                ui.label("5. 案件アラート作成").style(
                    "font-size: 18px; font-weight: bold; margin-bottom: 8px; width: 380px; margin: 0 auto;")
                make_button("【一括】全工程を実行",
                            lambda: (
                                create_alert.make_folders(),
                                create_alert.create_sheets(),
                                create_alert.check_sheets_at_creation(),
                                create_alert.make_send_list()
                            ),
                            log, primary=True, task_name="案件アラート作成 全工程")
                make_button("フォルダ作成",       create_alert.make_folders,              log, task_name="案件アラート作成 > フォルダ作成")
                make_button("アラート作成",       create_alert.create_sheets,             log, task_name="案件アラート作成 > アラート作成")
                make_button("作成後チェック",     create_alert.check_sheets_at_creation,  log, task_name="案件アラート作成 > 作成後チェック")
                make_button("送付対象一覧作成",   create_alert.make_send_list,            log, task_name="案件アラート作成 > 送付対象一覧作成")

ui.run(
    title="サンレモ成約捕捉 処理ツール",
    reload=False,
    native=True,
    window_size=(1440, 810)
)
