from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from core.csv_reader import read_csv
from core.data_writer import write_dataset
from core.demo_assets import DemoAssets, build_demo_assets
from core.doc_converter import ConversionResult, libreoffice_exists, prepare_template
from core.excel_reader import ExcelDataset, read_excel
from core.settings_store import AppSettings, SettingsStore
from core.template_engine import TagStatus, build_tag_statuses, extract_tags, merge_documents
from ui.data_panel import DataPanel
from ui.tag_panel import TagPanel

THEME_PRESETS = {
    "Obsidian Violet": {
        "accent": "#6d4aff",
        "accent_hover": "#5735e8",
        "success": "#2b8a78",
        "success_hover": "#216b5d",
        "danger": "#c2416c",
        "danger_hover": "#a53359",
        "muted": "#9aa4b2",
    },
    "Obsidian Emerald": {
        "accent": "#1f9d84",
        "accent_hover": "#187a67",
        "success": "#3cb371",
        "success_hover": "#2f8a58",
        "danger": "#c2416c",
        "danger_hover": "#a53359",
        "muted": "#9fb2aa",
    },
    "Obsidian Amber": {
        "accent": "#d97706",
        "accent_hover": "#b65f04",
        "success": "#2b8a78",
        "success_hover": "#216b5d",
        "danger": "#b45309",
        "danger_hover": "#92400e",
        "muted": "#b4a78d",
    },
    "Graphite Blue": {
        "accent": "#2563eb",
        "accent_hover": "#1d4ed8",
        "success": "#0f766e",
        "success_hover": "#115e59",
        "danger": "#be123c",
        "danger_hover": "#9f1239",
        "muted": "#94a3b8",
    },
    "Rose Night": {
        "accent": "#e11d48",
        "accent_hover": "#be123c",
        "success": "#2b8a78",
        "success_hover": "#216b5d",
        "danger": "#9d174d",
        "danger_hover": "#831843",
        "muted": "#b6a0ad",
    },
}


class WordMergeApp:
    def __init__(self, root: ctk.CTk) -> None:
        self.root = root
        self.root.title("Word 版型合併工具 v0.0.1")
        self.root.geometry("1380x860")
        self.root.minsize(1180, 760)
        self.root.configure(fg_color="#101114")

        self.excel_path: Path | None = None
        self.template_path: Path | None = None
        self.output_dir: Path = Path(__file__).resolve().parents[1] / "output"
        self.dataset: ExcelDataset | None = None
        self.template_tags: list[str] = []
        self.demo_assets: DemoAssets | None = None
        self.settings_store = SettingsStore(Path(__file__).resolve().parents[1] / "settings.json")
        self.settings = self.settings_store.load()
        self.data_dir_var = ctk.StringVar(value="")
        self.template_dir_var = ctk.StringVar(value="")
        self.output_dir_var = ctk.StringVar(value=str(self.output_dir))
        self.data_file_var = ctk.StringVar(value="")
        self.template_file_var = ctk.StringVar(value="")
        self.naming_field_var = ctk.StringVar(value="")
        self.theme_var = ctk.StringVar(value=self.settings.theme or "Obsidian Violet")
        self.font_scale_var = ctk.IntVar(value=self.settings.font_scale or 100)
        self.status_var = ctk.StringVar(value="Obsidian Mode｜等待資料載入")
        self.footer_var = ctk.StringVar(value="已選 0 筆資料｜版型：未選擇")
        self.case_var = ctk.StringVar(value="案例：尚未載入")
        self.paths_var = ctk.StringVar(value="Excel：未選擇｜版型：未選擇")

        self._build_layout()
        self._load_demo_assets()
        self.refresh_footer()

    def _build_layout(self) -> None:
        shell = ctk.CTkFrame(self.root, fg_color="#101114", corner_radius=0)
        shell.pack(fill="both", expand=True)

        title_row = ctk.CTkFrame(shell, fg_color="#101114")
        title_row.pack(fill="x", padx=20, pady=(18, 10))
        title_left = ctk.CTkFrame(title_row, fg_color="transparent")
        title_left.pack(side="left")
        ctk.CTkLabel(
            title_left,
            text="Word 版型合併工具",
            text_color="#f8fafc",
            font=ctk.CTkFont(family="Helvetica", size=30, weight="bold"),
        ).pack(anchor="w")
        ctk.CTkLabel(
            title_left,
            text="Obsidian Mode｜深色批次合併介面，預載真實商務案例",
            text_color="#9aa4b2",
            font=ctk.CTkFont(size=13),
        ).pack(anchor="w", pady=(4, 0))
        ctk.CTkLabel(
            title_row,
            textvariable=self.status_var,
            text_color="#a78bfa",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="right")

        self.page_tabs = ctk.CTkTabview(
            shell,
            fg_color="#101114",
            segmented_button_fg_color="#151821",
            segmented_button_selected_color="#6d4aff",
            segmented_button_selected_hover_color="#5735e8",
            segmented_button_unselected_color="#232833",
            segmented_button_unselected_hover_color="#2f3745",
            text_color="#e5e7eb",
        )
        self.page_tabs.pack(fill="both", expand=True, padx=20, pady=(0, 12))
        self.page_tabs._segmented_button.configure(  # type: ignore[attr-defined]
            height=42,
            font=ctk.CTkFont(size=16, weight="bold"),
        )
        self.page_tabs.add("主工作區")
        self.page_tabs.add("設定")
        self.page_tabs.add("說明")

        workspace_shell = self.page_tabs.tab("主工作區")
        workspace_shell.configure(fg_color="#101114")

        toolbar = ctk.CTkFrame(workspace_shell, fg_color="#151821", corner_radius=20, border_width=1, border_color="#2f3440")
        toolbar.pack(fill="x", pady=(0, 12))
        ctk.CTkLabel(
            toolbar,
            text="主工作區",
            text_color="#dbe4f0",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left", padx=12, pady=12)
        ctk.CTkLabel(
            toolbar,
            text="設定與說明已獨立成完整頁面",
            text_color="#9aa4b2",
            font=ctk.CTkFont(size=13),
        ).pack(side="left", padx=8)
        ctk.CTkButton(
            toolbar,
            text="批次產出",
            command=self.generate_documents,
            fg_color="#c2416c",
            hover_color="#a53359",
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="right", padx=12, pady=12)

        options = ctk.CTkFrame(workspace_shell, fg_color="#101114")
        options.pack(fill="x", pady=(0, 12))
        ctk.CTkLabel(
            options,
            text="檔名欄位",
            text_color="#dbe4f0",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left")
        self.naming_field_combo = ctk.CTkComboBox(
            options,
            variable=self.naming_field_var,
            values=[""],
            width=180,
            state="readonly",
            fg_color="#1a1d24",
            border_color="#323847",
            dropdown_fg_color="#151821",
            button_color="#6d4aff",
            button_hover_color="#5735e8",
            text_color="#eef2ff",
        )
        self.naming_field_combo.pack(side="left", padx=(10, 18))
        ctk.CTkButton(
            options,
            text="全選資料",
            command=self.select_all_rows,
            fg_color="#232833",
            text_color="#e5e7eb",
            hover_color="#2f3745",
            height=38,
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            options,
            text="清除選取",
            command=self.clear_selected_rows,
            fg_color="#232833",
            text_color="#e5e7eb",
            hover_color="#2f3745",
            height=38,
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left")

        workspace = ctk.CTkFrame(workspace_shell, fg_color="#101114")
        workspace.pack(fill="both", expand=True)
        workspace.grid_columnconfigure(0, weight=3)
        workspace.grid_columnconfigure(1, weight=2)
        workspace.grid_columnconfigure(1, minsize=460)
        workspace.grid_rowconfigure(0, weight=1)

        self.data_panel = DataPanel(
            workspace,
            self.on_selection_changed,
            self.save_current_row,
            self.save_source_file,
            self.update_single_cell,
            self.save_headers,
        )
        self.data_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=10)

        self.tag_panel = TagPanel(workspace)
        self.tag_panel.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=10)

        settings_page = self.page_tabs.tab("設定")
        settings_page.configure(fg_color="#101114")
        settings_scroll = ctk.CTkScrollableFrame(settings_page, fg_color="#101114", corner_radius=0)
        settings_scroll.pack(fill="both", expand=True)
        self._build_settings_content(settings_scroll)

        info_page = self.page_tabs.tab("說明")
        info_page.configure(fg_color="#101114")
        info_scroll = ctk.CTkScrollableFrame(info_page, fg_color="#101114", corner_radius=0)
        info_scroll.pack(fill="both", expand=True)
        self._build_info_content(info_scroll)

        footer = ctk.CTkFrame(shell, fg_color="#151821", corner_radius=18, border_width=1, border_color="#2f3440")
        footer.pack(fill="x", padx=20, pady=(0, 20))
        ctk.CTkLabel(
            footer,
            textvariable=self.footer_var,
            text_color="#dbe4f0",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).pack(side="left", padx=14, pady=12)
        self.shell = shell
        self.toolbar = toolbar
        self.footer = footer
        self.workspace_shell = workspace_shell
        self.title_status_label = title_row.winfo_children()[-1]
        self.generate_button = toolbar.winfo_children()[-1]
        self.settings_content_root = settings_page
        self.info_content_root = info_page

    def _build_template_picker(self, parent) -> None:
        header = ctk.CTkFrame(parent, fg_color="#1f2229")
        header.pack(fill="x", padx=16, pady=(16, 10))
        self.template_picker_summary_var = ctk.StringVar(value="版型清單")
        ctk.CTkLabel(
            header,
            text="版型選擇",
            text_color="#f8fafc",
            font=ctk.CTkFont(size=22, weight="bold"),
        ).pack(side="left")
        ctk.CTkLabel(
            header,
            textvariable=self.template_picker_summary_var,
            text_color="#8f9bad",
            font=ctk.CTkFont(size=14),
        ).pack(side="right")

        body = ctk.CTkFrame(parent, fg_color="#1f2229")
        body.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=2)
        body.grid_rowconfigure(0, weight=1)

        self.template_checklist = ctk.CTkScrollableFrame(body, fg_color="#151821", corner_radius=16)
        self.template_checklist.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        preview_card = ctk.CTkFrame(body, fg_color="#151821", corner_radius=16)
        preview_card.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        ctk.CTkLabel(
            preview_card,
            text="內容預覽",
            text_color="#dbe4f0",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(anchor="w", padx=14, pady=(14, 8))
        self.template_preview = ctk.CTkTextbox(
            preview_card,
            fg_color="#11131a",
            text_color="#cbd5e1",
            border_color="#2f3440",
            border_width=1,
            wrap="word",
            font=ctk.CTkFont(size=15),
        )
        self.template_preview.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        self.template_preview.insert("1.0", "請從左側勾選一份版型。")
        self.template_preview.configure(state="disabled")
        self.template_checkbox_vars = {}

    def _build_settings_content(self, parent) -> None:
        card = ctk.CTkFrame(parent, fg_color="#151821", corner_radius=20, border_width=1, border_color="#2f3440")
        card.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(
            card,
            text="資料夾設定",
            text_color="#f8fafc",
            font=ctk.CTkFont(size=22, weight="bold"),
        ).pack(anchor="w", padx=18, pady=(18, 8))
        ctk.CTkLabel(
            card,
            text="這裡指定有資料的資料夾。工作台會直接從這些路徑列出 CSV/Excel 與 Word 範本。",
            text_color="#9aa4b2",
            font=ctk.CTkFont(size=13),
        ).pack(anchor="w", padx=18, pady=(0, 18))

        appearance = ctk.CTkFrame(card, fg_color="transparent")
        appearance.pack(fill="x", padx=18, pady=(0, 12))
        ctk.CTkLabel(
            appearance,
            text="主題",
            text_color="#dbe4f0",
            font=ctk.CTkFont(size=14, weight="bold"),
            width=140,
        ).pack(side="left")
        self.theme_combo = ctk.CTkComboBox(
            appearance,
            variable=self.theme_var,
            values=list(THEME_PRESETS.keys()),
            width=260,
            state="readonly",
            command=lambda _value: self.apply_visual_settings(),
            fg_color="#1a1d24",
            border_color="#323847",
            dropdown_fg_color="#151821",
            button_color="#6d4aff",
            button_hover_color="#5735e8",
            text_color="#eef2ff",
        )
        self.theme_combo.pack(side="left", padx=(10, 18))
        ctk.CTkLabel(
            appearance,
            text="字型倍率",
            text_color="#dbe4f0",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left")
        self.font_scale_value = ctk.StringVar(value=f"{self.font_scale_var.get()}%")
        self.font_scale_slider = ctk.CTkSlider(
            appearance,
            from_=90,
            to=130,
            number_of_steps=8,
            variable=self.font_scale_var,
            command=self.on_font_scale_changed,
            width=180,
            button_color="#6d4aff",
            button_hover_color="#5735e8",
            progress_color="#6d4aff",
        )
        self.font_scale_slider.pack(side="left", padx=(10, 10))
        ctk.CTkLabel(
            appearance,
            textvariable=self.font_scale_value,
            text_color="#eef2ff",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).pack(side="left")

        self._build_path_row(card, "CSV / Excel 資料夾", self.data_dir_var, self.choose_data_dir)
        self._build_path_row(card, "Word 範本資料夾", self.template_dir_var, self.choose_template_dir)
        self._build_path_row(card, "輸出資料夾", self.output_dir_var, self.choose_output_dir, is_output=True)

        file_row = ctk.CTkFrame(card, fg_color="transparent")
        file_row.pack(fill="x", padx=18, pady=(12, 8))
        ctk.CTkLabel(file_row, text="資料檔", text_color="#dbe4f0", font=ctk.CTkFont(size=14, weight="bold")).pack(side="left")
        self.data_file_combo = ctk.CTkComboBox(
            file_row,
            variable=self.data_file_var,
            values=[""],
            width=260,
            state="readonly",
            command=self.on_data_file_selected,
            fg_color="#1a1d24",
            border_color="#323847",
            dropdown_fg_color="#151821",
            button_color="#6d4aff",
            button_hover_color="#5735e8",
            text_color="#eef2ff",
        )
        self.data_file_combo.pack(side="left", padx=(12, 18))
        ctk.CTkLabel(file_row, text="Word 範本", text_color="#dbe4f0", font=ctk.CTkFont(size=14, weight="bold")).pack(side="left")
        self.template_file_combo = ctk.CTkComboBox(
            file_row,
            variable=self.template_file_var,
            values=[""],
            width=280,
            state="readonly",
            command=self.on_template_file_selected,
            fg_color="#1a1d24",
            border_color="#323847",
            dropdown_fg_color="#151821",
            button_color="#2b8a78",
            button_hover_color="#216b5d",
            text_color="#eef2ff",
        )
        self.template_file_combo.pack(side="left", padx=(12, 0))

        template_picker_wrap = ctk.CTkFrame(card, fg_color="transparent")
        template_picker_wrap.pack(fill="both", expand=True, padx=18, pady=(18, 10))
        self.template_picker_card = ctk.CTkFrame(
            template_picker_wrap,
            fg_color="#1f2229",
            corner_radius=20,
            border_width=1,
            border_color="#2f3440",
        )
        self.template_picker_card.pack(fill="both", expand=True)
        self._build_template_picker(self.template_picker_card)

        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.pack(fill="x", padx=18, pady=(18, 18))
        ctk.CTkButton(
            actions,
            text="重新掃描資料夾",
            command=self.refresh_folder_sources,
            fg_color="#232833",
            text_color="#e5e7eb",
            hover_color="#2f3745",
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left")
        ctk.CTkButton(
            actions,
            text="儲存設定",
            command=self.save_settings,
            fg_color="#6d4aff",
            hover_color="#5735e8",
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left", padx=10)
        ctk.CTkButton(
            actions,
            text="套用設定",
            command=self.apply_settings_and_refresh,
            fg_color="#c2416c",
            hover_color="#a53359",
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left")

    def _build_info_content(self, shell) -> None:
        ctk.CTkLabel(
            shell,
            text="說明",
            text_color="#f8fafc",
            font=ctk.CTkFont(size=28, weight="bold"),
        ).pack(anchor="w", padx=20, pady=(20, 8))
        ctk.CTkLabel(
            shell,
            text="獨立完整說明頁。工作台只保留操作內容。",
            text_color="#9aa4b2",
            font=ctk.CTkFont(size=13),
        ).pack(anchor="w", padx=20, pady=(0, 12))
        card = ctk.CTkFrame(shell, fg_color="#151821", corner_radius=20, border_width=1, border_color="#2f3440")
        card.pack(fill="both", expand=True, padx=20, pady=20)
        ctk.CTkLabel(
            card,
            text="顧問合約與付款通知批次產出",
            text_color="#eef2ff",
            font=ctk.CTkFont(size=24, weight="bold"),
        ).pack(anchor="w", padx=18, pady=(18, 10))
        ctk.CTkLabel(
            card,
            textvariable=self.case_var,
            text_color="#dce7f8",
            font=ctk.CTkFont(size=16, weight="bold"),
            justify="left",
            wraplength=860,
        ).pack(anchor="w", padx=18, pady=(0, 12))
        ctk.CTkLabel(
            card,
            text=(
                "範例資料包含公司名稱、專案名稱、付款期限、聯絡窗口與合約編號。\n"
                "可直接測試 CSV/Excel 對應、Word Tag 掃描、欄位編輯、存回來源檔與批次輸出。"
            ),
            text_color="#aeb8c7",
            font=ctk.CTkFont(size=15),
            justify="left",
            wraplength=860,
        ).pack(anchor="w", padx=18, pady=(0, 12))
        ctk.CTkLabel(
            card,
            textvariable=self.paths_var,
            text_color="#93a4bb",
            font=ctk.CTkFont(size=14),
            justify="left",
            wraplength=860,
        ).pack(anchor="w", padx=18, pady=(0, 18))

    def _build_path_row(self, parent, label: str, variable, command, is_output: bool = False) -> None:
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=18, pady=8)
        ctk.CTkLabel(row, text=label, text_color="#dbe4f0", font=ctk.CTkFont(size=14, weight="bold"), width=140).pack(side="left")
        ctk.CTkEntry(
            row,
            textvariable=variable,
            fg_color="#1a1d24",
            border_color="#323847",
            text_color="#eef2ff",
            height=38,
        ).pack(side="left", fill="x", expand=True, padx=(10, 10))
        btn_text = "選擇輸出資料夾" if is_output else "選擇資料夾"
        ctk.CTkButton(
            row,
            text=btn_text,
            command=command,
            fg_color="#232833",
            text_color="#e5e7eb",
            hover_color="#2f3745",
            height=38,
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left")

    def _load_demo_assets(self) -> None:
        project_root = Path(__file__).resolve().parents[1]
        self.demo_assets = build_demo_assets(project_root)
        if self.settings.output_dir:
            self.output_dir = Path(self.settings.output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir_var.set(str(self.output_dir))
        default_data_dir = Path(self.settings.data_dir) if self.settings.data_dir else self.demo_assets.csv_path.parent
        default_template_dir = Path(self.settings.template_dir) if self.settings.template_dir else self.demo_assets.template_path.parent
        self.data_dir_var.set(str(default_data_dir))
        self.template_dir_var.set(str(default_template_dir))
        self.case_var.set(
            "案例：三筆顧問合作資料，已附一份 CSV、Excel，以及四份不同風格的 Word 範本。設定頁可改成任何有資料的資料夾。"
        )
        self.apply_visual_settings()
        self.refresh_folder_sources()
        self.status_var.set("Obsidian Mode｜已載入示範案例資料夾")

    def on_font_scale_changed(self, _value=None) -> None:
        self.font_scale_value.set(f"{int(self.font_scale_var.get())}%")
        self.apply_visual_settings()

    def apply_visual_settings(self) -> None:
        theme_name = self.theme_var.get() if self.theme_var.get() in THEME_PRESETS else "Obsidian Violet"
        preset = THEME_PRESETS[theme_name]
        scale = max(0.9, min(1.3, int(self.font_scale_var.get()) / 100))
        ctk.set_widget_scaling(scale)
        ctk.set_window_scaling(scale)

        self.root.configure(fg_color="#101114")
        self.shell.configure(fg_color="#101114")
        self.workspace_shell.configure(fg_color="#101114")
        self.page_tabs.configure(
            segmented_button_selected_color=preset["accent"],
            segmented_button_selected_hover_color=preset["accent_hover"],
        )
        self.toolbar.configure(fg_color="#151821")
        self.footer.configure(fg_color="#151821")
        self.title_status_label.configure(text_color=preset["accent"])
        self.generate_button.configure(fg_color=preset["danger"], hover_color=preset["danger_hover"])
        if hasattr(self, "theme_combo"):
            self.theme_combo.configure(button_color=preset["accent"], button_hover_color=preset["accent_hover"])
        if hasattr(self, "font_scale_slider"):
            self.font_scale_slider.configure(button_color=preset["accent"], button_hover_color=preset["accent_hover"], progress_color=preset["accent"])

    def _tabular_files(self, directory: Path) -> list[Path]:
        if not directory.exists():
            return []
        return sorted(
            [
                path
                for path in directory.iterdir()
                if path.is_file() and path.suffix.lower() in {".csv", ".xlsx", ".xls"}
            ]
        )

    def _template_files(self, directory: Path) -> list[Path]:
        if not directory.exists():
            return []
        return sorted(
            [
                path
                for path in directory.iterdir()
                if path.is_file() and path.suffix.lower() in {".docx", ".doc"}
            ]
        )

    def refresh_folder_sources(self) -> None:
        data_dir = Path(self.data_dir_var.get()).expanduser() if self.data_dir_var.get().strip() else None
        template_dir = Path(self.template_dir_var.get()).expanduser() if self.template_dir_var.get().strip() else None
        data_files = self._tabular_files(data_dir) if data_dir else []
        template_files = self._template_files(template_dir) if template_dir else []

        self.data_files_map = {path.name: path for path in data_files}
        self.template_files_map = {path.name: path for path in template_files}
        if hasattr(self, "data_file_combo"):
            self.data_file_combo.configure(values=list(self.data_files_map.keys()) or [""])
        if hasattr(self, "template_file_combo"):
            self.template_file_combo.configure(values=list(self.template_files_map.keys()) or [""])
        self._refresh_template_picker()

        if data_files:
            selected_name = self.data_file_var.get() if self.data_file_var.get() in self.data_files_map else data_files[0].name
            self.data_file_var.set(selected_name)
            self._load_dataset(self.data_files_map[selected_name])
        else:
            self.data_file_var.set("")

        if template_files:
            selected_name = self.template_file_var.get() if self.template_file_var.get() in self.template_files_map else template_files[0].name
            self.template_file_var.set(selected_name)
            self._load_template(self.template_files_map[selected_name])
        else:
            self.template_file_var.set("")
            self._refresh_template_picker()

        self.refresh_footer()

    def choose_data_dir(self) -> None:
        path = filedialog.askdirectory(title="選擇 CSV / Excel 資料夾")
        if not path:
            return
        self.data_dir_var.set(path)
        self.refresh_folder_sources()
        self.status_var.set("已更新資料檔資料夾")

    def choose_template_dir(self) -> None:
        path = filedialog.askdirectory(title="選擇 Word 範本資料夾")
        if not path:
            return
        self.template_dir_var.set(path)
        self.refresh_folder_sources()
        self.status_var.set("已更新範本資料夾")

    def save_settings(self) -> None:
        self.output_dir = Path(self.output_dir_var.get()).expanduser()
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.settings = AppSettings(
            data_dir=self.data_dir_var.get().strip(),
            template_dir=self.template_dir_var.get().strip(),
            output_dir=str(self.output_dir),
            theme=self.theme_var.get().strip(),
            font_scale=int(self.font_scale_var.get()),
        )
        self.settings_store.save(self.settings)
        self.apply_visual_settings()
        self.status_var.set("設定已儲存")

    def apply_settings_and_refresh(self) -> None:
        self.save_settings()
        self.refresh_folder_sources()
        if hasattr(self, "settings_window") and self.settings_window.winfo_exists():
            self.settings_window.focus()

    def _load_dataset(self, path: Path) -> None:
        if path.suffix.lower() == ".csv":
            self.dataset = read_csv(path)
        else:
            self.dataset = read_excel(path)
        self.excel_path = path
        self.data_panel.load_rows(self.dataset.headers, self.dataset.rows)
        self.naming_field_combo.configure(values=[""] + self.dataset.headers)
        self.naming_field_var.set("合約編號" if "合約編號" in self.dataset.headers else "")
        self.paths_var.set(
            f"資料：{self.excel_path.name}\n版型：{self.template_path.name if self.template_path else '未選擇'}\n資料夾：{self.data_dir_var.get() or '未設定'}"
        )

    def _load_template(self, path: Path) -> None:
        self.template_path = path
        conversion: ConversionResult = prepare_template(self.template_path)
        try:
            self.template_tags = extract_tags(conversion.converted_path)
            preview_text = self._extract_template_preview_text(conversion.converted_path)
        finally:
            conversion.cleanup()
        self.paths_var.set(
            f"資料：{self.excel_path.name if self.excel_path else '未選擇'}\n版型：{self.template_path.name}\n範本夾：{self.template_dir_var.get() or '未設定'}"
        )
        self._set_template_preview(preview_text)
        self._refresh_template_picker()
        self.refresh_tag_preview()

    def _extract_template_preview_text(self, template_path: Path) -> str:
        try:
            from core.template_engine import extract_template_preview

            return extract_template_preview(template_path)
        except Exception:
            return "無法讀取版型內容預覽。"

    def _set_template_preview(self, text: str) -> None:
        self.template_preview.configure(state="normal")
        self.template_preview.delete("1.0", "end")
        self.template_preview.insert("1.0", text or "版型內容為空白。")
        self.template_preview.configure(state="disabled")

    def _refresh_template_picker(self) -> None:
        if not hasattr(self, "template_checklist"):
            return
        for child in self.template_checklist.winfo_children():
            child.destroy()
        self.template_checkbox_vars = {}
        names = list(getattr(self, "template_files_map", {}).keys())
        self.template_picker_summary_var.set(f"共 {len(names)} 份版型，可勾選切換")
        for name in names:
            var = ctk.BooleanVar(value=self.template_path is not None and self.template_path.name == name)
            self.template_checkbox_vars[name] = var
            checkbox = ctk.CTkCheckBox(
                self.template_checklist,
                text=name,
                variable=var,
                onvalue=True,
                offvalue=False,
                command=lambda selected=name: self.on_template_checked(selected),
                text_color="#e5e7eb",
                font=ctk.CTkFont(size=15, weight="bold"),
                fg_color="#6d4aff",
                hover_color="#5735e8",
                border_color="#596273",
            )
            checkbox.pack(anchor="w", fill="x", padx=12, pady=8)

    def on_template_checked(self, selected_name: str) -> None:
        for name, var in self.template_checkbox_vars.items():
            var.set(name == selected_name)
        self.on_template_file_selected(selected_name)

    def on_data_file_selected(self, selected_name: str) -> None:
        path = getattr(self, "data_files_map", {}).get(selected_name)
        if not path:
            return
        try:
            self._load_dataset(path)
        except Exception as exc:
            messagebox.showerror("資料讀取失敗", str(exc))
            return
        self.status_var.set(f"已切換資料：{path.name}")
        self.refresh_footer()

    def on_template_file_selected(self, selected_name: str) -> None:
        path = getattr(self, "template_files_map", {}).get(selected_name)
        if not path:
            return
        try:
            self._load_template(path)
        except Exception as exc:
            messagebox.showerror("版型讀取失敗", str(exc))
            return
        self.status_var.set(f"已切換版型：{path.name}")
        self.refresh_footer()

    def import_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="選擇 Excel 檔案",
            filetypes=[("Excel", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            self._load_dataset(Path(path))
        except Exception as exc:
            messagebox.showerror("Excel 匯入失敗", str(exc))
            return

        self.status_var.set(f"已匯入 {self.excel_path.name}")
        self.refresh_tag_preview()
        self.refresh_footer()

    def choose_template(self) -> None:
        path = filedialog.askopenfilename(
            title="選擇 Word 版型",
            filetypes=[("Word", "*.docx *.doc"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            self._load_template(Path(path))
        except RuntimeError as exc:
            if Path(path).suffix.lower() == ".doc" and not libreoffice_exists():
                messagebox.showwarning("LibreOffice 未安裝", f"{exc}\n\n請安裝 LibreOffice 後再使用 .doc 版型。")
            else:
                messagebox.showerror("版型讀取失敗", str(exc))
            return
        except Exception as exc:
            messagebox.showerror("版型讀取失敗", str(exc))
            return

        if not self.template_tags:
            messagebox.showwarning("版型無 Tag", "版型中找不到 {{ }} Tag")
        self.status_var.set(f"已選擇 {self.template_path.name}")
        self.refresh_tag_preview()
        self.refresh_footer()

    def choose_output_dir(self) -> None:
        path = filedialog.askdirectory(title="選擇輸出資料夾")
        if not path:
            return
        self.output_dir = Path(path)
        self.output_dir_var.set(path)
        self.status_var.set(f"輸出資料夾：{self.output_dir}")
        self.refresh_footer()

    def on_selection_changed(self, _event=None) -> None:
        self._sync_editor_with_selection()
        self.refresh_footer()

    def select_all_rows(self) -> None:
        self.data_panel.select_all()
        self.refresh_footer()

    def clear_selected_rows(self) -> None:
        self.data_panel.clear_selection()
        self.refresh_footer()

    def refresh_tag_preview(self) -> None:
        if not self.dataset and not self.template_tags:
            self.tag_panel.clear("尚無 Tag")
            return

        headers = self.dataset.headers if self.dataset else []
        sample_row = None
        if self.dataset and self.dataset.rows:
            selected_index = self.data_panel.get_primary_selected_index()
            if selected_index is not None and 0 <= selected_index < len(self.dataset.rows):
                sample_row = self.dataset.rows[selected_index]
            else:
                sample_row = self.dataset.rows[0]
        statuses = build_tag_statuses(self.template_tags, headers, sample_row)
        self.tag_panel.render(statuses)

    def _selected_rows(self) -> list[dict[str, str]]:
        if not self.dataset:
            return []
        indices = self.data_panel.get_selected_indices()
        if not indices:
            return self.dataset.rows
        return [self.dataset.rows[index] for index in indices if 0 <= index < len(self.dataset.rows)]

    def _sync_editor_with_selection(self) -> None:
        if not self.dataset:
            self.data_panel.load_editor([], {}, None)
            return
        index = self.data_panel.get_primary_selected_index()
        if index is None or index >= len(self.dataset.rows):
            self.data_panel.load_editor(self.dataset.headers, {}, None)
            return
        self.data_panel.load_editor(self.dataset.headers, self.dataset.rows[index], index)

    def save_current_row(self, index: int, payload: dict[str, str]) -> None:
        if not self.dataset or not (0 <= index < len(self.dataset.rows)):
            return
        for header in self.dataset.headers:
            self.dataset.rows[index][header] = payload.get(header, "")
        self.data_panel.update_row(index, self.dataset.headers, self.dataset.rows[index])
        self._sync_editor_with_selection()
        self.status_var.set(f"已更新第 {index + 1} 筆資料，尚未寫回來源檔")

    def update_single_cell(self, row_index: int, header: str, value: str) -> None:
        if not self.dataset or not (0 <= row_index < len(self.dataset.rows)):
            return
        self.dataset.rows[row_index][header] = value
        self.data_panel.update_row(row_index, self.dataset.headers, self.dataset.rows[row_index])
        selected_index = self.data_panel.get_primary_selected_index()
        if selected_index == row_index:
            self._sync_editor_with_selection()
        self.status_var.set(f"已更新第 {row_index + 1} 筆的 {header}，尚未寫回來源檔")

    def save_source_file(self, index: int, payload: dict[str, str]) -> None:
        self.save_current_row(index, payload)
        if not self.dataset or not self.excel_path:
            return
        try:
            write_dataset(self.excel_path, self.dataset.headers, self.dataset.rows)
        except Exception as exc:
            messagebox.showerror("存檔失敗", str(exc))
            return
        self.status_var.set(f"已存回來源檔：{self.excel_path.name}")

    def save_headers(self, payload: list[tuple[str, str]]) -> None:
        if not self.dataset:
            return

        new_headers: list[str] = []
        seen: set[str] = set()
        rename_map: dict[str, str] = {}
        for old_header, new_header in payload:
            final_header = new_header or old_header
            if not final_header:
                messagebox.showwarning("欄名無效", "欄名不能留空")
                return
            if final_header in seen:
                messagebox.showwarning("欄名重複", f"欄名重複：{final_header}")
                return
            seen.add(final_header)
            new_headers.append(final_header)
            rename_map[old_header] = final_header

        self.dataset.rows = [
            {rename_map[old]: row.get(old, "") for old in self.dataset.headers}
            for row in self.dataset.rows
        ]
        self.dataset.headers = new_headers
        self.data_panel.load_rows(self.dataset.headers, self.dataset.rows)
        self.naming_field_combo.configure(values=[""] + self.dataset.headers)
        if self.naming_field_var.get() not in self.dataset.headers:
            self.naming_field_var.set("合約編號" if "合約編號" in self.dataset.headers else "")
        self.refresh_tag_preview()

        if self.excel_path:
            try:
                write_dataset(self.excel_path, self.dataset.headers, self.dataset.rows)
            except Exception as exc:
                messagebox.showerror("欄名存檔失敗", str(exc))
                return

        self.status_var.set("欄名已更新並存回來源檔")
        self.refresh_footer()

    def _missing_tag_statuses(self) -> list[TagStatus]:
        if not self.dataset:
            return []
        statuses = build_tag_statuses(
            self.template_tags,
            self.dataset.headers,
            self.dataset.rows[0] if self.dataset.rows else None,
        )
        return [item for item in statuses if item.status == "missing"]

    def generate_documents(self) -> None:
        if not self.dataset:
            messagebox.showwarning("缺少資料", "請先匯入 Excel")
            return
        if not self.template_path:
            messagebox.showwarning("缺少版型", "請先選擇 Word 版型")
            return

        missing = self._missing_tag_statuses()
        if missing:
            proceed = messagebox.askyesno(
                "Tag 對應不完整",
                "有部分 Tag 找不到對應欄位，產出時會留空。是否仍要繼續？",
            )
            if not proceed:
                return

        try:
            conversion = prepare_template(self.template_path)
        except Exception as exc:
            messagebox.showerror("版型準備失敗", str(exc))
            return

        try:
            summary = merge_documents(
                template_path=conversion.converted_path,
                rows=self._selected_rows(),
                output_dir=self.output_dir,
                naming_field=self.naming_field_var.get().strip(),
            )
        except PermissionError:
            messagebox.showerror("輸出失敗", "輸出資料夾無寫入權限，請更換輸出路徑")
            return
        except Exception as exc:
            messagebox.showerror("批次產出失敗", str(exc))
            return
        finally:
            conversion.cleanup()

        message_lines = [
            f"✅ 成功產出：{summary.success_count} 份",
            f"⚠️ 有缺失欄位：{summary.warning_count} 份",
            f"❌ 失敗：{summary.failure_count} 份",
            "",
            f"輸出位置：{self.output_dir}",
        ]
        if summary.warnings:
            message_lines.append("")
            message_lines.extend(f"- 第 {warning.row_index} 筆：{warning.message}" for warning in summary.warnings[:5])
        if summary.failures:
            message_lines.append("")
            message_lines.extend(f"- 第 {failure.row_index} 筆失敗：{failure.reason}" for failure in summary.failures[:5])
        messagebox.showinfo("批次完成", "\n".join(message_lines))
        self._open_output_folder()
        self.status_var.set(f"已完成批次產出，共 {summary.success_count} 份")
        self.refresh_footer()

    def _open_output_folder(self) -> None:
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(self.output_dir))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", str(self.output_dir)], check=False)
            else:
                subprocess.run(["xdg-open", str(self.output_dir)], check=False)
        except Exception:
            messagebox.showwarning("開啟資料夾失敗", f"文件已產出，但無法自動開啟資料夾：{self.output_dir}")

    def refresh_footer(self) -> None:
        selected_count = len(self.data_panel.get_selected_indices()) if self.dataset else 0
        template_name = self.template_path.name if self.template_path else "未選擇"
        row_count = len(self.dataset.rows) if self.dataset else 0
        display_count = selected_count or row_count
        self.footer_var.set(f"已選 {display_count} 筆資料｜版型：{template_name}｜輸出：{self.output_dir}")


def main() -> None:
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    root = ctk.CTk()
    WordMergeApp(root)
    root.mainloop()
