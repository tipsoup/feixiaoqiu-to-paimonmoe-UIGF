#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Simple Tkinter GUI for feixiaoqiu-to-paimonmoe / UIGF conversion."""

from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

import converter
import to_uigf


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("非小酋抽卡记录转换器")
        self.root.geometry("720x520")

        self.input_path = tk.StringVar()
        self.uid = tk.StringVar()
        self.use_heuristic = tk.BooleanVar(value=True)
        self.uigf_version = tk.StringVar(value="v4.2")
        self.queue = queue.Queue()

        self._build_ui()
        self.root.after(100, self._poll_queue)

    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill="both", expand=True)

        # Input file
        row1 = ttk.Frame(frame)
        row1.pack(fill="x", **pad)
        ttk.Label(row1, text="非小酋 Excel:").pack(side="left")
        ttk.Entry(row1, textvariable=self.input_path).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(row1, text="选择文件", command=self._browse_input).pack(side="left")

        # UID
        row2 = ttk.Frame(frame)
        row2.pack(fill="x", **pad)
        ttk.Label(row2, text="UID（仅 UIGF 需要）:").pack(side="left")
        ttk.Entry(row2, textvariable=self.uid).pack(side="left", fill="x", expand=True, padx=5)

        # Options
        row3 = ttk.Frame(frame)
        row3.pack(fill="x", **pad)
        ttk.Checkbutton(row3, text="启用双卡池 Wish 2 启发式识别", variable=self.use_heuristic).pack(side="left")
        ttk.Label(row3, text="    UIGF 版本:").pack(side="left")
        ttk.Combobox(
            row3,
            textvariable=self.uigf_version,
            values=["v3.0", "v4.2"],
            state="readonly",
            width=8,
        ).pack(side="left")

        # Buttons
        row4 = ttk.Frame(frame)
        row4.pack(fill="x", **pad)
        self.btn_paimon = ttk.Button(row4, text="转换为 Paimon.moe Excel", command=self._convert_paimon)
        self.btn_paimon.pack(side="left", padx=5)
        self.btn_uigf = ttk.Button(row4, text="转换为 UIGF JSON", command=self._convert_uigf)
        self.btn_uigf.pack(side="left", padx=5)

        # Log
        self.log = scrolledtext.ScrolledText(frame, height=18, state="disabled")
        self.log.pack(fill="both", expand=True, **pad)

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="选择非小酋导出的 Excel",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if path:
            self.input_path.set(path)

    def _log(self, msg: str):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _set_busy(self, busy: bool):
        state = "disabled" if busy else "normal"
        self.btn_paimon.configure(state=state)
        self.btn_uigf.configure(state=state)

    def _run(self, fn):
        self._set_busy(True)
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        threading.Thread(target=lambda: self._worker(fn), daemon=True).start()

    def _worker(self, fn):
        try:
            result = fn()
            self.queue.put(("ok", result))
        except Exception as exc:
            self.queue.put(("error", repr(exc)))

    def _poll_queue(self):
        try:
            kind, payload = self.queue.get_nowait()
        except queue.Empty:
            pass
        else:
            self._set_busy(False)
            if kind == "ok":
                warnings = payload
                if warnings:
                    self._log("转换完成，但有警告：")
                    for w in warnings:
                        self._log(" - " + w)
                else:
                    self._log("转换完成，无警告。")
                messagebox.showinfo("完成", "转换完成！")
            else:
                self._log("出错了：" + payload)
                messagebox.showerror("错误", payload)
        self.root.after(100, self._poll_queue)

    def _get_input(self) -> Path | None:
        path = self.input_path.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择非小酋 Excel 文件。")
            return None
        p = Path(path)
        if not p.is_file():
            messagebox.showerror("错误", "输入文件不存在。")
            return None
        return p

    def _convert_paimon(self):
        src = self._get_input()
        if src is None:
            return
        default_out = src.with_name(src.stem + "_paimonmoe.xlsx")
        out = filedialog.asksaveasfilename(
            title="保存 Paimon.moe Excel",
            defaultextension=".xlsx",
            initialfile=default_out.name,
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not out:
            return
        self._log(f"读取: {src}")
        self._run(lambda: converter.convert_paimon_file(
            src,
            out,
            use_heuristic=self.use_heuristic.get(),
        ))

    def _convert_uigf(self):
        src = self._get_input()
        if src is None:
            return
        uid = self.uid.get().strip()
        if not uid:
            messagebox.showwarning("提示", "UIGF 导出需要填写 UID。")
            return
        default_out = src.with_name(src.stem + "_uigf.json")
        out = filedialog.asksaveasfilename(
            title="保存 UIGF JSON",
            defaultextension=".json",
            initialfile=default_out.name,
            filetypes=[("JSON 文件", "*.json")],
        )
        if not out:
            return
        self._log(f"读取: {src}")
        self._run(lambda: to_uigf.convert_to_uigf_file(
            src,
            out,
            uid,
            use_heuristic=self.use_heuristic.get(),
            uigf_version=self.uigf_version.get(),
        ))


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
