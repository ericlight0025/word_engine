# 架構文件 — Word 版型合併工具

## 專案概覽

本工具是一套 **桌面 GUI 應用程式**，用於將 CSV / Excel 資料表批次合併進 Word 版型（.docx / .doc），並輸出獨立文件。採用 Python + CustomTkinter，在本機執行，無需伺服器。

---

## 目錄結構

```
word_engine/
├── main.py                  # 入口點
├── core/                    # 業務邏輯層（無 UI 相依）
│   ├── csv_reader.py        # CSV 讀取
│   ├── excel_reader.py      # Excel 讀取
│   ├── data_writer.py       # CSV / Excel 回寫
│   ├── doc_converter.py     # .doc → .docx 轉換（LibreOffice）
│   ├── template_engine.py   # Tag 解析、狀態對應、批次合併
│   ├── settings_store.py    # 設定讀寫（JSON）
│   └── demo_assets.py       # 示範資料產生器
├── ui/                      # UI 元件層
│   ├── app.py               # 主控制器（WordMergeApp）
│   ├── data_panel.py        # 資料表格 + 編輯面板
│   └── tag_panel.py         # Tag 對應狀態面板
├── tests/
│   └── test_core.py         # core 層單元測試
├── assets/demo/             # 示範 CSV、Excel、Word 範本
└── settings.json            # 執行期使用者設定（自動產生）
```

---

## 三層架構

```
┌─────────────────────────────────────────┐
│                 ui/app.py               │  ← 控制器：協調所有模組
│             WordMergeApp                │
└───────┬────────────────────┬────────────┘
        │                    │
┌───────▼──────┐    ┌────────▼───────────┐
│  ui/          │    │  core/             │
│  data_panel   │    │  csv_reader        │
│  tag_panel    │    │  excel_reader      │
└───────────────┘    │  data_writer       │
                     │  doc_converter     │
                     │  template_engine   │
                     │  settings_store    │
                     │  demo_assets       │
                     └────────────────────┘
```

- **core 層**：純業務邏輯，不依賴任何 UI 框架，可獨立測試。
- **ui 層**：CustomTkinter 元件，只負責呈現與事件轉發。
- **app.py**：唯一的橋樑，持有應用狀態（`dataset`、`template_path`…），呼叫 core 並更新 UI。

---

## 模組說明

### core/excel_reader.py

讀取 `.xlsx` / `.xls`，回傳統一的 `ExcelDataset`。

```
ExcelDataset
├── headers: list[str]
└── rows:    list[dict[str, str]]
```

- 使用 **openpyxl**（read-only 模式），讀完即 close。
- 驗證：空內容、空欄名、重複欄名均 raise `ValueError`。

### core/csv_reader.py

讀取 `.csv`，同樣回傳 `ExcelDataset`。依賴 `excel_reader.ExcelDataset`（共用型別）。

### core/data_writer.py

`write_dataset(path, headers, rows)` — 自動判斷副檔名，以 openpyxl 寫 `.xlsx`，以 `csv.DictWriter` 寫 `.csv`（UTF-8 BOM）。

### core/doc_converter.py

```
prepare_template(path) -> ConversionResult
```

- `.docx`：直接回傳原路徑，`cleanup()` 為 no-op。
- `.doc`：呼叫 `soffice --headless --convert-to docx`，回傳暫存目錄路徑，`cleanup()` 會刪除暫存目錄。

呼叫端必須在 `finally` 內呼叫 `conversion.cleanup()`，app.py 的 `_load_template` 與 `generate_documents` 均已遵守。

### core/template_engine.py

核心引擎，三個主要職責：

| 函式 | 說明 |
|------|------|
| `extract_tags(path)` | 用 zipfile 解析 `word/document.xml`、header/footer XML，正規式抓 `{{ tag }}`，回傳已排序去重的 `list[str]` |
| `extract_template_preview(path)` | 同樣解析 document.xml，剝除 XML 標籤後回傳前 2000 字純文字 |
| `build_tag_statuses(tags, headers, sample_row)` | 對比版型 Tag 與資料欄位，產生三種狀態：`matched` / `missing` / `extra` |
| `merge_documents(...)` | 批次合併：一次讀入版型 bytes，對每列建立 `DocxTemplate(BytesIO(bytes))`，`render(row)`，`save(target)` |

**錯誤處理策略**

- 逐列 `except OSError: raise`：磁碟滿、無寫入權限等系統性錯誤立即上拋，不繼續執行。
- 其他例外（如渲染失敗）：記錄進 `MergeFailure`，繼續處理下一列。
- 回傳 `MergeSummary`，呼叫端（app.py）決定如何顯示。

```
MergeSummary
├── success_count:  int
├── warning_count:  int
├── failure_count:  int
├── output_files:   list[Path]
├── warnings:       list[MergeWarning]   # 如命名欄位為空
└── failures:       list[MergeFailure]   # 如渲染例外
```

### core/settings_store.py

以 `settings.json` 持久化 `AppSettings`（data_dir、template_dir、output_dir、theme、font_scale）。`load()` 內任何解析錯誤均靜默回傳預設值，不 crash 應用。

### core/demo_assets.py

`build_demo_assets(root)` 在 `assets/demo/` 產生示範 CSV、Excel 與四份 Word 範本（幂等：已存在則跳過）。應用啟動時呼叫一次。

### ui/app.py — WordMergeApp

應用程式唯一的狀態持有者與控制器：

| 狀態屬性 | 型別 | 說明 |
|----------|------|------|
| `dataset` | `ExcelDataset \| None` | 目前載入的資料 |
| `excel_path` | `Path \| None` | 來源檔路徑（供回寫用） |
| `template_path` | `Path \| None` | 目前版型路徑（僅於成功載入後設定） |
| `template_tags` | `list[str]` | 版型內所有 Tag |
| `output_dir` | `Path` | 輸出資料夾 |
| `settings` | `AppSettings` | 目前生效設定 |

主要方法群組：

- **載入**：`_load_dataset()`、`_load_template()`、`_load_demo_assets()`
- **資料夾掃描**：`refresh_folder_sources()`
- **編輯**：`save_current_row()`、`update_single_cell()`、`save_source_file()`、`save_headers()`
- **Tag 預覽**：`refresh_tag_preview()`
- **批次產出**：`generate_documents()`
- **UI 構建**：`_build_layout()`、`_build_settings_content()`、`_build_info_content()`

### ui/data_panel.py — DataPanel

`ctk.CTkFrame` 子類，包含三個區域：

1. **欄名客製化區**（`CTkScrollableFrame`）：顯示可編輯的欄名，送出後呼叫 `on_save_headers` callback。
2. **資料表格**（`ttk.Treeview`）：最多顯示 `MAX_VISIBLE_COLUMNS`（預設 6）欄，支援多選、排序、雙擊 inline 編輯。
3. **完整欄位編輯區**（`CTkScrollableFrame`）：顯示選取列的所有欄位，送出後呼叫 `on_save_row` / `on_save_source` callback。

所有資料異動均透過 callback 向上傳至 `app.py`，`DataPanel` 本身不持有業務狀態。

### ui/tag_panel.py — TagPanel

`ctk.CTkFrame` 子類，接收 `list[TagStatus]` 並以顏色卡片呈現：

| 狀態 | 顏色語義 |
|------|---------|
| `matched` | 綠色 — Tag 有對應欄位 |
| `missing` | 紅色 — 版型 Tag 找不到欄位 |
| `extra` | 黃色 — 欄位有但版型未用到 |

---

## 完整資料流

### 啟動

```
main.py
  └─ _validate_tk_version()
  └─ ui.app.main()
       └─ WordMergeApp.__init__()
            ├─ SettingsStore.load()
            ├─ _build_layout()
            ├─ _load_demo_assets()
            │    └─ build_demo_assets(root)
            │    └─ refresh_folder_sources()
            │         ├─ _load_dataset(csv/xlsx)
            │         └─ _load_template(docx)
            └─ refresh_footer()
```

### 批次產出

```
使用者按「批次產出」
  └─ generate_documents()
       ├─ _missing_tag_statuses()  ← 確認 Tag 對應
       ├─ _selected_rows()         ← 無選取 = 全部
       ├─ prepare_template(template_path)
       └─ merge_documents(converted_path, rows, output_dir, naming_field)
            ├─ source.read_bytes()             ← 讀一次版型
            └─ for each row:
                 ├─ DocxTemplate(BytesIO(bytes))
                 ├─ render(dict(row))
                 └─ save(output_dir / filename)
       └─ 顯示 MergeSummary 結果訊息框
       └─ _open_output_folder()
```

### 資料回寫

```
使用者按「存回原始檔」
  └─ DataPanel._save_source()
  └─ app.save_source_file(index, payload)
       ├─ save_current_row()  ← 先更新記憶體
       └─ write_dataset(excel_path, headers, rows)  ← 覆寫原檔
```

---

## 相依套件

| 套件 | 用途 |
|------|------|
| `customtkinter` | 深色主題 GUI 框架 |
| `openpyxl` | Excel 讀寫 |
| `docxtpl` | Word 版型渲染（基於 Jinja2） |
| `python-docx` | docxtpl 底層依賴 |

外部工具：**LibreOffice**（可選），僅在處理 `.doc` 格式時需要。

---

## 測試

```
tests/test_core.py
```

涵蓋 core 層四個場景：

| 測試 | 驗證內容 |
|------|---------|
| `test_read_excel_skips_blank_rows` | 空白列不進入 rows |
| `test_extract_tags_reads_document_and_header` | 同時掃 document + header XML |
| `test_build_tag_statuses_marks_missing_and_extra` | matched / missing / extra 三態正確 |
| `test_merge_documents_uses_naming_field_and_fallback` | 命名欄位有值用欄位值，空值用流水號 |

執行：

```bash
python -m pytest tests/
```
