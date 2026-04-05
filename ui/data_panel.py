from __future__ import annotations

import tkinter as tk
from tkinter import ttk

import customtkinter as ctk


class DataPanel(ctk.CTkFrame):
    MAX_VISIBLE_COLUMNS = 6

    def __init__(
        self,
        master: tk.Misc,
        on_selection_changed,
        on_save_row,
        on_save_source,
        on_cell_updated,
        on_save_headers,
    ) -> None:
        super().__init__(master, fg_color="#1f2229", corner_radius=20, border_width=1, border_color="#2f3440")
        self.on_selection_changed = on_selection_changed
        self.on_save_row = on_save_row
        self.on_save_source = on_save_source
        self.on_cell_updated = on_cell_updated
        self.on_save_headers = on_save_headers
        self.editor_entries: dict[str, ctk.CTkEntry] = {}
        self.header_entries: list[tuple[str, ctk.CTkEntry]] = []
        self.headers: list[str] = []
        self.visible_headers: list[str] = []
        self.rows_data: list[dict[str, str]] = []
        self.sort_state: dict[str, bool] = {}
        self.inline_entry: tk.Entry | None = None
        self.inline_item_id: str | None = None
        self.inline_column_index: int | None = None

        header = ctk.CTkFrame(self, fg_color="#1f2229")
        header.pack(fill="x", padx=16, pady=(16, 10))
        self.summary_var = tk.StringVar(value="尚未匯入 Excel")
        ctk.CTkLabel(
            header,
            text="Excel 資料區",
            text_color="#f1f5f9",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(side="left")
        ctk.CTkLabel(
            header,
            textvariable=self.summary_var,
            text_color="#8f9bad",
            font=ctk.CTkFont(size=12),
        ).pack(side="right")
        self.hint_var = tk.StringVar(value="雙擊表格儲存格可直接修改；超過 6 欄請用下方編輯區")
        ctk.CTkLabel(
            self,
            textvariable=self.hint_var,
            text_color="#a8b3c7",
            font=ctk.CTkFont(size=13),
            anchor="w",
        ).pack(fill="x", padx=16, pady=(0, 10))

        header_card = ctk.CTkFrame(self, fg_color="#151821", corner_radius=16)
        header_card.pack(fill="x", padx=16, pady=(0, 12))
        header_top = ctk.CTkFrame(header_card, fg_color="#151821")
        header_top.pack(fill="x", padx=14, pady=(14, 10))
        self.header_title_var = tk.StringVar(value="欄名客製化")
        ctk.CTkLabel(
            header_top,
            textvariable=self.header_title_var,
            text_color="#f1f5f9",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(side="left")
        ctk.CTkButton(
            header_top,
            text="儲存欄名",
            command=self._save_headers,
            fg_color="#d97706",
            hover_color="#b65f04",
            height=36,
            font=ctk.CTkFont(size=13, weight="bold"),
        ).pack(side="right")
        self.header_scroll = ctk.CTkScrollableFrame(header_card, fg_color="#11131a", corner_radius=14, height=110)
        self.header_scroll.pack(fill="x", padx=14, pady=(0, 14))

        table_wrap = ctk.CTkFrame(self, fg_color="#151821", corner_radius=16)
        table_wrap.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "WordMerge.Treeview",
            background="#151821",
            fieldbackground="#151821",
            foreground="#e5e7eb",
            rowheight=30,
            borderwidth=0,
        )
        style.configure(
            "WordMerge.Treeview.Heading",
            background="#232833",
            foreground="#d8dee9",
            relief="flat",
            font=("Helvetica", 11, "bold"),
        )
        style.map(
            "WordMerge.Treeview",
            background=[("hover", "#151821"), ("!selected", "#151821"), ("selected", "#6d4aff")],
            fieldbackground=[("hover", "#151821"), ("!selected", "#151821"), ("selected", "#6d4aff")],
            foreground=[("hover", "#e5e7eb"), ("!selected", "#e5e7eb"), ("selected", "#f8fafc")],
        )
        style.map(
            "WordMerge.Treeview.Heading",
            background=[("active", "#232833")],
            foreground=[("active", "#d8dee9")],
        )

        self.table = ttk.Treeview(table_wrap, show="headings", selectmode="extended", style="WordMerge.Treeview")
        self.table.grid(row=0, column=0, sticky="nsew", padx=(12, 0), pady=(12, 0))
        self.table.bind("<<TreeviewSelect>>", self.on_selection_changed)
        self.table.bind("<Double-1>", self._begin_inline_edit)

        scroll_y = ctk.CTkScrollbar(table_wrap, orientation="vertical", command=self.table.yview)
        scroll_y.grid(row=0, column=1, sticky="ns", padx=(10, 12), pady=(12, 0))
        scroll_x = ctk.CTkScrollbar(table_wrap, orientation="horizontal", command=self.table.xview)
        scroll_x.grid(row=1, column=0, sticky="ew", padx=(12, 0), pady=(10, 12))
        self.table.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        table_wrap.rowconfigure(0, weight=1)
        table_wrap.columnconfigure(0, weight=1)

        editor_card = ctk.CTkFrame(self, fg_color="#151821", corner_radius=16)
        editor_card.pack(fill="x", padx=16, pady=(0, 16))
        editor_header = ctk.CTkFrame(editor_card, fg_color="#151821")
        editor_header.pack(fill="x", padx=14, pady=(14, 10))
        self.editor_title_var = tk.StringVar(value="完整欄位編輯｜請先選擇一筆資料")
        ctk.CTkLabel(
            editor_header,
            textvariable=self.editor_title_var,
            text_color="#f1f5f9",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(side="left")
        self.editor_hint_var = tk.StringVar(value="前 6 欄可直接雙擊表格修改；所有欄位都可在這裡完整編輯並存檔")
        ctk.CTkLabel(
            editor_card,
            textvariable=self.editor_hint_var,
            text_color="#a8b3c7",
            font=ctk.CTkFont(size=13),
            anchor="w",
        ).pack(fill="x", padx=14, pady=(0, 10))

        self.editor_scroll = ctk.CTkScrollableFrame(editor_card, fg_color="#11131a", corner_radius=14, height=220)
        self.editor_scroll.pack(fill="x", padx=14, pady=(0, 10))

        actions = ctk.CTkFrame(editor_card, fg_color="#151821")
        actions.pack(fill="x", padx=14, pady=(0, 14))
        ctk.CTkButton(
            actions,
            text="存回表格",
            command=self._save_row,
            fg_color="#6d4aff",
            hover_color="#5735e8",
            height=38,
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left")
        ctk.CTkButton(
            actions,
            text="存回原始檔",
            command=self._save_source,
            fg_color="#2b8a78",
            hover_color="#216b5d",
            height=38,
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left", padx=10)

    def load_rows(self, headers: list[str], rows: list[dict[str, str]]) -> None:
        self.headers = headers
        self.rows_data = rows
        self.visible_headers = headers[: self.MAX_VISIBLE_COLUMNS]
        self.table.delete(*self.table.get_children())
        table_columns = list(self.visible_headers)
        has_extra_columns = len(headers) > self.MAX_VISIBLE_COLUMNS
        if has_extra_columns:
            table_columns.append("__more__")
        self.table["columns"] = table_columns

        for column in self.visible_headers:
            width = max(120, min(240, len(column) * 18 + 36))
            self.table.heading(column, text=column, command=lambda selected=column: self.sort_by_column(selected))
            self.table.column(column, width=width, stretch=True, anchor="w")
        if has_extra_columns:
            self.table.heading("__more__", text="更多欄位")
            self.table.column("__more__", width=180, stretch=False, anchor="center")

        for index, row in enumerate(rows, start=1):
            values = [row.get(header, "") for header in self.visible_headers]
            if has_extra_columns:
                values.append(f"其餘 {len(headers) - self.MAX_VISIBLE_COLUMNS} 欄請下方編輯")
            self.table.insert("", "end", iid=str(index), values=values)

        if rows:
            self.table.selection_set(self.table.get_children())
        if has_extra_columns:
            self.summary_var.set(f"共 {len(rows)} 筆，表格最多顯示 6 欄；其餘欄位請用下方編輯區")
        else:
            self.summary_var.set(f"共 {len(rows)} 筆，已預設全選，可局部產出")
        self.hint_var.set(
            "雙擊表格儲存格可直接修改；按 Enter 套用。超過 6 欄請用下方編輯區後存回原始檔"
        )
        self.load_header_editor(headers)
        self.load_editor(headers, rows[0] if rows else {}, 0 if rows else None)

    def get_selected_indices(self) -> list[int]:
        return [int(item) - 1 for item in self.table.selection()]

    def get_primary_selected_index(self) -> int | None:
        indices = self.get_selected_indices()
        return indices[0] if indices else None

    def select_all(self) -> None:
        self.table.selection_set(self.table.get_children())

    def clear_selection(self) -> None:
        self.table.selection_remove(self.table.selection())

    def load_header_editor(self, headers: list[str]) -> None:
        for child in self.header_scroll.winfo_children():
            child.destroy()
        self.header_entries = []

        for header in headers:
            line = ctk.CTkFrame(self.header_scroll, fg_color="transparent")
            line.pack(fill="x", padx=10, pady=6)
            ctk.CTkLabel(
                line,
                text=header,
                text_color="#dbe4f0",
                font=ctk.CTkFont(size=13, weight="bold"),
                width=140,
                anchor="w",
            ).pack(side="left")
            entry = ctk.CTkEntry(
                line,
                fg_color="#1a1d24",
                border_color="#323847",
                text_color="#eef2ff",
            )
            entry.pack(side="left", fill="x", expand=True, padx=(10, 0))
            entry.insert(0, header)
            self.header_entries.append((header, entry))

    def _save_headers(self) -> None:
        payload = [(old, entry.get().strip()) for old, entry in self.header_entries]
        self.on_save_headers(payload)

    def load_editor(self, headers: list[str], row: dict[str, str], index: int | None) -> None:
        for child in self.editor_scroll.winfo_children():
            child.destroy()
        self.editor_entries = {}

        if index is None:
            self.editor_title_var.set("完整欄位編輯｜請先選擇一筆資料")
            return

        self.editor_title_var.set(f"完整欄位編輯｜第 {index + 1} 筆")
        for header in headers:
            line = ctk.CTkFrame(self.editor_scroll, fg_color="transparent")
            line.pack(fill="x", padx=10, pady=6)
            ctk.CTkLabel(
                line,
                text=header,
                text_color="#dbe4f0",
                font=ctk.CTkFont(size=13, weight="bold"),
                width=120,
                anchor="w",
            ).pack(side="left")
            entry = ctk.CTkEntry(
                line,
                fg_color="#1a1d24",
                border_color="#323847",
                text_color="#eef2ff",
            )
            entry.pack(side="left", fill="x", expand=True, padx=(10, 0))
            entry.insert(0, row.get(header, ""))
            self.editor_entries[header] = entry

    def _editor_payload(self) -> dict[str, str]:
        return {header: entry.get() for header, entry in self.editor_entries.items()}

    def _save_row(self) -> None:
        index = self.get_primary_selected_index()
        if index is None:
            return
        self.on_save_row(index, self._editor_payload())

    def _save_source(self) -> None:
        index = self.get_primary_selected_index()
        if index is None:
            return
        self.on_save_source(index, self._editor_payload())

    def update_row(self, index: int, headers: list[str], row: dict[str, str]) -> None:
        item_id = str(index + 1)
        if self.table.exists(item_id):
            values = [row.get(header, "") for header in self.visible_headers]
            if len(headers) > self.MAX_VISIBLE_COLUMNS:
                values.append(f"其餘 {len(headers) - self.MAX_VISIBLE_COLUMNS} 欄請下方編輯")
            self.table.item(item_id, values=values)

    def sort_by_column(self, column: str) -> None:
        descending = self.sort_state.get(column, False)
        self.sort_state[column] = not descending
        children = list(self.table.get_children())

        def sort_key(item_id: str):
            row_index = int(item_id) - 1
            value = self.rows_data[row_index].get(column, "")
            text = str(value).strip()
            try:
                return (0, float(text.replace(",", "")))
            except ValueError:
                return (1, text.casefold())

        ordered = sorted(children, key=sort_key, reverse=descending)
        for new_position, item_id in enumerate(ordered):
            self.table.move(item_id, "", new_position)
        direction = "↓" if descending else "↑"
        for visible in self.visible_headers:
            label = f"{visible} {direction}" if visible == column else visible
            self.table.heading(visible, text=label, command=lambda selected=visible: self.sort_by_column(selected))

    def _begin_inline_edit(self, event) -> None:
        region = self.table.identify("region", event.x, event.y)
        if region != "cell":
            return
        item_id = self.table.identify_row(event.y)
        column_id = self.table.identify_column(event.x)
        if not item_id or not column_id:
            return

        column_index = int(column_id.replace("#", "")) - 1
        if column_index < 0 or column_index >= len(self.visible_headers):
            return

        self._close_inline_editor(save=False)
        bbox = self.table.bbox(item_id, column_id)
        if not bbox:
            return
        x, y, width, height = bbox
        current_values = list(self.table.item(item_id, "values"))
        current_text = current_values[column_index] if column_index < len(current_values) else ""

        entry = tk.Entry(
            self.table,
            bg="#111827",
            fg="#f8fafc",
            insertbackground="#f8fafc",
            relief="solid",
            highlightthickness=1,
        )
        entry.insert(0, current_text)
        entry.select_range(0, "end")
        entry.focus()
        entry.place(x=x, y=y, width=width, height=height)
        entry.bind("<Return>", lambda _event: self._close_inline_editor(save=True))
        entry.bind("<Escape>", lambda _event: self._close_inline_editor(save=False))
        entry.bind("<FocusOut>", lambda _event: self._close_inline_editor(save=True))

        self.inline_entry = entry
        self.inline_item_id = item_id
        self.inline_column_index = column_index

    def _close_inline_editor(self, save: bool) -> None:
        if self.inline_entry is None:
            return

        entry = self.inline_entry
        item_id = self.inline_item_id
        column_index = self.inline_column_index
        new_value = entry.get()
        entry.destroy()
        self.inline_entry = None
        self.inline_item_id = None
        self.inline_column_index = None

        if not save or item_id is None or column_index is None:
            return

        row_index = int(item_id) - 1
        header = self.visible_headers[column_index]
        self.on_cell_updated(row_index, header, new_value)
