import sys
import tkinter
from tkinter import messagebox

from ui.app import main


def _validate_tk_version() -> None:
    if tkinter.TkVersion >= 8.6:
        return

    message = (
        "目前執行環境的 Tk 版本過舊，偵測到 Tk "
        f"{tkinter.TkVersion:.1f}。\n\n"
        "customtkinter 需要較新的 Tk 8.6+，否則可能出現整窗全黑。\n\n"
        f"Python: {sys.executable}\n"
        "請改用 python.org 或 Homebrew 安裝的 Python 後再啟動。"
    )
    try:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror("Tk 版本過舊", message)
        root.destroy()
    except Exception:
        print(message, file=sys.stderr)
    raise SystemExit(1)


if __name__ == "__main__":
    _validate_tk_version()
    main()
