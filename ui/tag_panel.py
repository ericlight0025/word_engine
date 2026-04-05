from __future__ import annotations

import tkinter as tk

import customtkinter as ctk

from core.template_engine import TagStatus


class TagPanel(ctk.CTkFrame):
    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master, fg_color="#1f2229", corner_radius=20, border_width=1, border_color="#2f3440")

        header = ctk.CTkFrame(self, fg_color="#1f2229")
        header.pack(fill="x", padx=16, pady=(16, 10))
        self.summary_var = tk.StringVar(value="尚未選擇版型")
        ctk.CTkLabel(
            header,
            text="目前資料套版預覽",
            text_color="#f1f5f9",
            font=ctk.CTkFont(size=22, weight="bold"),
        ).pack(side="left")
        ctk.CTkLabel(
            header,
            textvariable=self.summary_var,
            text_color="#8f9bad",
            font=ctk.CTkFont(size=14),
        ).pack(side="right")

        self.scrollable = ctk.CTkScrollableFrame(
            self,
            fg_color="#151821",
            corner_radius=16,
        )
        self.scrollable.pack(fill="both", expand=True, padx=16, pady=(0, 16))

    def render(self, statuses: list[TagStatus]) -> None:
        for child in self.scrollable.winfo_children():
            child.destroy()

        palette = {
            "matched": ("#13261f", "#7ee0a1", "✅"),
            "missing": ("#302612", "#f4cb67", "⚠️"),
            "extra": ("#32161b", "#ff8c94", "❌"),
        }

        matched = sum(1 for item in statuses if item.status == "matched")
        total = len([item for item in statuses if item.status != "extra"])
        self.summary_var.set(f"已對應 {matched}/{total} 個版型 Tag")

        for item in statuses:
            bg, fg, icon = palette.get(item.status, ("#1d2330", "#cbd5e1", "•"))
            card = ctk.CTkFrame(self.scrollable, fg_color=bg, corner_radius=14, border_width=1, border_color="#303848")
            card.pack(fill="x", pady=6, padx=2)
            ctk.CTkLabel(
                card,
                text=f"{icon} {{{{ {item.tag} }}}}",
                text_color=fg,
                font=ctk.CTkFont(size=18, weight="bold"),
                anchor="w",
            ).pack(fill="x", padx=14, pady=(12, 4))
            ctk.CTkLabel(
                card,
                text=item.message,
                text_color=fg,
                font=ctk.CTkFont(size=15),
                anchor="w",
                justify="left",
                wraplength=420,
            ).pack(fill="x", padx=14, pady=(0, 12))

    def clear(self, message: str) -> None:
        self.render([TagStatus(tag=message, status="missing", message="請先匯入資料並選擇版型")])
