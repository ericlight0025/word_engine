from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core.excel_reader import ExcelDataset
from ui.app import WordMergeApp
from ui.data_panel import DataPanel


class Value:
    """提供無顯示器測試所需的變數行為。"""
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class UiSafetyTests(unittest.TestCase):
    def make_app(self):
        app = object.__new__(WordMergeApp)
        app.root = Mock()
        app.dataset = ExcelDataset(["姓名"], [{"姓名": "原始"}])
        app.excel_path = Path("source.csv")
        app.template_path = Path("template.docx")
        app.template_tags = ["姓名"]
        app.has_unsaved_changes = False
        app.data_panel = Mock(editor_row_index=None)
        app.data_panel.get_selected_indices.return_value = []
        app.data_panel.get_primary_selected_index.return_value = None
        app.naming_field_combo = Mock()
        app.data_file_combo = Mock()
        app.template_file_combo = Mock()
        app.output_dir = Path("out")
        for name in ("status", "footer", "paths", "data_file", "template_file", "naming_field", "data_dir", "template_dir"):
            setattr(app, f"{name}_var", Value())
        app.refresh_tag_preview = Mock()
        app._sync_editor_with_selection = Mock()
        app._set_template_preview = Mock()
        app._refresh_template_picker = Mock()
        return app

    def test_cleared_selection_never_generates_all_rows(self):
        app = self.make_app()
        self.assertEqual(app._selected_rows(), [])
        with patch("ui.app.messagebox.showwarning") as warning, patch("ui.app.prepare_template") as prepare:
            app.generate_documents()
        warning.assert_called_once_with("未選取資料", "請先選取要產出的資料列；清除選取後不會產出文件。")
        prepare.assert_not_called()

    def test_footer_reports_zero_when_selection_is_empty(self):
        app = self.make_app()
        app.refresh_footer()
        self.assertTrue(app.footer_var.get().startswith("已選 0 筆"))

    def test_editor_pending_changes_survive_selection_switch(self):
        app = self.make_app()
        app.data_panel.editor_row_index = 0
        app.data_panel._editor_payload.return_value = {"姓名": "尚未按存回的修改"}
        app.on_selection_changed()
        self.assertEqual(app.dataset.rows[0]["姓名"], "尚未按存回的修改")
        self.assertTrue(app.has_unsaved_changes)
        app.refresh_tag_preview.assert_called_once()

    def test_cancel_unsaved_close_keeps_window_and_data(self):
        app = self.make_app()
        app.has_unsaved_changes = True
        with patch("ui.app.messagebox.askyesnocancel", return_value=None):
            app._on_close()
        app.root.destroy.assert_not_called()
        self.assertEqual(app.dataset.rows[0]["姓名"], "原始")

    def test_failed_save_before_close_does_not_close_or_clear_dirty_flag(self):
        app = self.make_app()
        app.has_unsaved_changes = True
        with (
            patch("ui.app.messagebox.askyesnocancel", return_value=True),
            patch("ui.app.write_dataset", side_effect=PermissionError("locked")),
            patch("ui.app.messagebox.showerror") as error,
        ):
            app._on_close()
        app.root.destroy.assert_not_called()
        self.assertTrue(app.has_unsaved_changes)
        error.assert_called_once()

    def test_cancel_file_switch_does_not_read_or_replace_dataset(self):
        app = self.make_app()
        original = app.dataset
        app.has_unsaved_changes = True
        with patch("ui.app.messagebox.askyesnocancel", return_value=None), patch("ui.app.read_csv") as read:
            self.assertFalse(app._load_dataset(Path("new.csv")))
        read.assert_not_called()
        self.assertIs(app.dataset, original)
        self.assertEqual(app.excel_path, Path("source.csv"))

    def test_template_load_failure_keeps_previous_template_and_tags(self):
        app = self.make_app()
        with tempfile.TemporaryDirectory() as directory:
            invalid = Path(directory) / "invalid.docx"
            invalid.write_bytes(b"not a docx")
            with self.assertRaises(Exception):
                app._load_template(invalid)
        self.assertEqual(app.template_path, Path("template.docx"))
        self.assertEqual(app.template_tags, ["姓名"])

    def test_failed_header_save_keeps_original_model(self):
        app = self.make_app()
        original = app.dataset
        with patch("ui.app.write_dataset", side_effect=OSError("locked")), patch("ui.app.messagebox.showerror") as error:
            app.save_headers([("姓名", "新姓名")])
        self.assertIs(app.dataset, original)
        self.assertEqual(app.dataset.headers, ["姓名"])
        self.assertEqual(app.dataset.rows, [{"姓名": "原始"}])
        app.data_panel.load_rows.assert_not_called()
        error.assert_called_once()

    def test_empty_header_is_rejected(self):
        app = self.make_app()
        with patch("ui.app.write_dataset") as write, patch("ui.app.messagebox.showwarning") as warning:
            app.save_headers([("姓名", " ")])
        write.assert_not_called()
        warning.assert_called_once()

    def test_empty_folders_clear_stale_sources_and_previews(self):
        app = self.make_app()
        with tempfile.TemporaryDirectory() as directory:
            app.data_dir_var.set(directory)
            app.template_dir_var.set(directory)
            app.refresh_folder_sources()
        self.assertIsNone(app.dataset)
        self.assertIsNone(app.excel_path)
        self.assertIsNone(app.template_path)
        self.assertEqual(app.template_tags, [])
        app.data_panel.load_rows.assert_called_once_with([], [])
        self.assertIn("未選擇", app.paths_var.get())

    def test_folder_refresh_preserves_dirty_current_dataset(self):
        app = self.make_app()
        app.has_unsaved_changes = True
        original = app.dataset
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.csv"
            source.write_text("姓名\n新的內容\n", encoding="utf-8")
            app.excel_path = source
            app.data_dir_var.set(directory)
            app.template_dir_var.set(directory)
            with patch.object(app, "_load_dataset") as load:
                app.refresh_folder_sources()
        load.assert_not_called()
        self.assertIs(app.dataset, original)
        self.assertTrue(app.has_unsaved_changes)

    def test_inline_focusout_reentry_commits_once(self):
        panel = object.__new__(DataPanel)
        panel.inline_entry = Mock()
        panel.inline_entry.get.return_value = "修改"
        panel.inline_entry.destroy.side_effect = lambda: panel._close_inline_editor(save=True)
        panel.inline_item_id = "1"
        panel.inline_column_index = 0
        panel.visible_headers = ["姓名"]
        panel.on_cell_updated = Mock()
        panel._close_inline_editor(save=True)
        panel.on_cell_updated.assert_called_once_with(0, "姓名", "修改")
        self.assertIsNone(panel.inline_entry)


if __name__ == "__main__":
    unittest.main()
