#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
サッカー記録 整形ツール (GUI)
- フォルダ または ファイルを選んで「実行」を押すだけ。
- 元ファイルは変更せず、整形済みコピーを _processed フォルダに出力する。
- format_records.py の run() を呼び出す(同じフォルダに置くこと。exe化時は自動同梱)。
"""
import os, sys, threading, queue

try:
    import format_records as fr
except ImportError:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import format_records as fr

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

APP_TITLE = "Excelデータ加工"

class App:
    def __init__(self, root):
        self.root = root
        self.inputs = []
        self.q = queue.Queue()
        self.running = False
        self.out_dir = None
        root.title(APP_TITLE)
        root.geometry("680x520")

        tk.Label(root, text="フォルダ か Excelファイル を選んで「実行」を押してください。",
                 font=("Meiryo", 11)).pack(anchor="w", padx=10, pady=8)

        btns = tk.Frame(root); btns.pack(anchor="w", padx=10)
        tk.Button(btns, text="📁 フォルダを選ぶ", width=18, command=self.pick_folder).pack(side="left", padx=4)
        tk.Button(btns, text="📄 ファイルを選ぶ", width=18, command=self.pick_files).pack(side="left", padx=4)

        self.sel = tk.Label(root, text="（未選択）", fg="#555", anchor="w",
                            justify="left", wraplength=640)
        self.sel.pack(anchor="w", padx=10, pady=6)

        self.del_keireki = tk.BooleanVar(value=True)
        tk.Checkbutton(root, text="「経歴」シートを削除する", variable=self.del_keireki,
                       font=("Meiryo", 10)).pack(anchor="w", padx=10)

        self.run_btn = tk.Button(root, text="▶ 実行", width=20, bg="#2d7d46", fg="white",
                                 font=("Meiryo", 11, "bold"), command=self.start)
        self.run_btn.pack(anchor="w", padx=10, pady=8)

        tk.Label(root, text="進捗ログ:", anchor="w").pack(anchor="w", padx=10)
        self.log = scrolledtext.ScrolledText(root, height=16, font=("Consolas", 9))
        self.log.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.root.after(100, self.poll)

    def pick_folder(self):
        d = filedialog.askdirectory(title="整形するExcelが入ったフォルダを選択")
        if d:
            self.inputs = [d]
            self.sel.config(text="選択フォルダ: " + d)

    def pick_files(self):
        fs = filedialog.askopenfilenames(title="整形するExcelファイルを選択",
                                         filetypes=[("Excel", "*.xlsx")])
        if fs:
            self.inputs = list(fs)
            self.sel.config(text="選択ファイル: %d 件\n" % len(fs)
                            + "\n".join(os.path.basename(f) for f in fs))

    def emit(self, s):
        self.q.put(("log", s))

    def poll(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "log":
                    self.log.insert("end", payload + "\n"); self.log.see("end")
                elif kind == "done":
                    self.finish(*payload)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def start(self):
        if self.running:
            return
        if not self.inputs:
            messagebox.showwarning(APP_TITLE, "先にフォルダかファイルを選んでください。")
            return
        self.running = True
        self.run_btn.config(state="disabled", text="処理中…")
        self.log.delete("1.0", "end")
        threading.Thread(target=self.worker, args=(self.del_keireki.get(),), daemon=True).start()

    def worker(self, delete):
        try:
            out, res = fr.run(self.inputs, delete_keireki=delete, progress=self.emit)
            self.out_dir = out
            ok = sum(1 for _, s in res if s == "OK")
            ng = [n for n, s in res if s != "OK"]
            self.q.put(("done", (ok, len(res), ", ".join(ng))))
        except Exception as e:
            self.q.put(("log", "!! 予期せぬエラー: %s" % e))
            self.q.put(("done", (0, 0, str(e))))

    def finish(self, ok, total, ng):
        self.running = False
        self.run_btn.config(state="normal", text="▶ 実行")
        msg = "完了しました。\n成功: %d / %d ファイル" % (ok, total)
        if ng:
            msg += "\n失敗: " + ng
        msg += "\n\n出力先: _processed フォルダ\n（元ファイルは変更していません）"
        messagebox.showinfo(APP_TITLE, msg)
        if self.out_dir and os.path.isdir(self.out_dir):
            try:
                os.startfile(self.out_dir)
            except Exception:
                pass

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
