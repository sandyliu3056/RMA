# -*- coding: utf-8 -*-
"""
exe 啟動器。PyInstaller 以這個檔案為進入點（檔名用英文，打包比較不會出問題），
實際程式在 零件規劃平台.py。打包成 exe 後沒有主控台，所以錯誤寫到 exe 旁邊的 啟動錯誤.log 並跳視窗提示。
"""
import os
import sys
import traceback


def _base_dir():
    return os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))


def main():
    try:
        import 零件規劃平台 as app
        app.App().mainloop()
    except Exception:  # noqa
        msg = traceback.format_exc()
        try:
            with open(os.path.join(_base_dir(), "啟動錯誤.log"), "w", encoding="utf-8") as fh:
                fh.write(msg)
        except OSError:
            pass
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk(); root.withdraw()
            messagebox.showerror("售後零件規劃平台", "程式無法啟動，詳細原因已寫到 啟動錯誤.log：\n\n" + msg[-1500:])
        except Exception:  # noqa
            print(msg, file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
