#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""EXE ビルドスクリプト。EXEを作成.bat から呼ばれる。
   pyinstaller/openpyxl を入れ、soccer_format_app.pyw を単一EXE化し、
   'Excelデータ加工.exe' という名前で同フォルダに出力する。"""
import os, sys, subprocess, shutil, traceback

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)
logf = open(os.path.join(HERE, "build_log.txt"), "w", encoding="utf-8")
def log(*a):
    s = " ".join(str(x) for x in a)
    print(s)
    logf.write(s + "\n"); logf.flush()

APP        = "soccer_format_app.pyw"
ASCII_NAME = "SoccerRecordFormatter"
OUT_NAME   = "Excelデータ加工.exe"

def main():
    try:
        log("Python:", sys.version)
        if not os.path.exists(APP):
            log("[ERROR] %s が同じフォルダにありません。" % APP); return 1
        if not os.path.exists("format_records.py"):
            log("[ERROR] format_records.py が同じフォルダにありません。"); return 1
        log("\n[1/3] 必要な部品をインストール中（初回は数分）...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pip"])
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pyinstaller", "openpyxl"])
        log("\n[2/3] EXE をビルド中...")
        import PyInstaller.__main__ as pim
        pim.run([
            "--onefile", "--windowed",
            "--name", ASCII_NAME,
            "--hidden-import", "openpyxl.cell._writer",
            "--noconfirm", "--clean",
            APP,
        ])
        src = os.path.join(HERE, "dist", ASCII_NAME + ".exe")
        if not os.path.exists(src):
            log("[ERROR] ビルド出力が見つかりません:", src); return 1
        log("\n[3/3] 出力名を変更中...")
        shutil.copyfile(src, os.path.join(HERE, OUT_NAME))
        log("\n==============================================")
        log("  完成: %s" % OUT_NAME)
        log("  このファイルを配布すれば、相手は何も入れずに")
        log("  ダブルクリックで使えます。")
        log("==============================================")
        return 0
    except subprocess.CalledProcessError as e:
        log("[ERROR] コマンド失敗:", e); traceback.print_exc(file=logf); return 1
    except Exception as e:
        log("[ERROR]", e); traceback.print_exc(file=logf); return 1

if __name__ == "__main__":
    sys.exit(main())
