# PPT 助手工具

一個功能強大的 Python PPT 生成工具,支持從 JSON 格式讀取內容,並輸出包含文字、圖片、表格、圖表的專業簡報。

**注意**: 此項目不再使用 Python 包裝系統,僅作為腳本分發。請使用 `pip install -r requirements.txt` 安裝依賴。

## 功能特色

- **JSON 格式輸入**: 使用結構化的 JSON 格式定義簡報內容
- **16:9 比例**: 自動生成標準 16:9 寬屏比例的簡報 (13.333 x 7.5 英寸)
- **多種佈局**: 支持標題頁、內容頁、雙欄佈局、圖片頁、表格頁、圖表頁
- **圖表類型**: 支持柱狀圖、條形圖、折線圖、圓餅圖
- **圖片插入**: 支持在簡報中插入圖片並添加說明
- **表格生成**: 自動生成格式化的表格
- **易於使用**: 簡單的命令行界面

## 安裝依賴

```bash
pip install -r requirements.txt
```

或使用 sudo (如果需要):

```bash
sudo pip install -r requirements.txt
```

## 使用方法

### 基本用法

```bash
python3 ppt_assistant.py <json_配置文件>
```

### 示例

```bash
python3 ppt_assistant.py example_config.json
```

## JSON 配置格式

### 基本結構

```json
{
  "output": "output.pptx",
  "footer": {
    "icon_path": "path/to/icon.png"
  },
  "slides": [
    {
      "layout": "佈局類型",
      "content": {
        // 內容配置
      }
    }
  ]
}
```

### 配置項說明

- **output**: (必需) 輸出文件路徑
- **footer**: (可選) 頁腳配置對象
  - **icon_path**: (可選) 圖標文件路徑,將顯示在頁腳左側 (支持 PNG, JPG, JPEG, GIF, BMP; SVG 需要額外依賴項)
  - 頁碼將自動顯示在右側 (格式: 當前頁/總頁數)
- **slides**: (必需) 投影片數組,每個元素定義一張投影片

**注意**: 生成的簡報會自動設置為 16:9 比例 (13.333 x 7.5 英寸)。

## 支持的佈局類型

### 1. 標題頁 (title)

用於簡報的首頁,包含主標題和副標題。

```json
{
  "layout": "title",
  "content": {
    "title": "簡報主標題",
    "subtitle": "副標題或說明文字"
  }
}
```

### 2. 內容頁 (content)

標準的內容頁,包含標題和正文內容。正文可以是字符串或列表。

```json
{
  "layout": "content",
  "content": {
    "title": "頁面標題",
    "body": "單段文字內容"
  }
}
```

或使用列表格式:

```json
{
  "layout": "content",
  "content": {
    "title": "頁面標題",
    "body": [
      "要點一",
      "要點二",
      "要點三"
    ]
  }
}
```

### 3. 雙欄佈局 (two_column)

將內容分為左右兩欄顯示。

```json
{
  "layout": "two_column",
  "content": {
    "title": "頁面標題",
    "left": "左欄內容\n可以使用換行符",
    "right": "右欄內容\n可以使用換行符"
  }
}
```

### 4. 圖片頁 (image)

插入圖片並可添加說明文字。

```json
{
  "layout": "image",
  "content": {
    "title": "頁面標題",
    "image": "path/to/image.png",
    "caption": "圖片說明文字(可選)"
  }
}
```

**注意**: 圖片路徑可以是相對路徑或絕對路徑。

### 5. 表格頁 (table)

生成格式化的表格。

```json
{
  "layout": "table",
  "content": {
    "title": "頁面標題",
    "table": {
      "headers": ["欄位1", "欄位2", "欄位3"],
      "rows": [
        ["數據1-1", "數據1-2", "數據1-3"],
        ["數據2-1", "數據2-2", "數據2-3"],
        ["數據3-1", "數據3-2", "數據3-3"]
      ]
    }
  }
}
```

### 6. 圖表頁 (chart)

生成各種類型的圖表。

```json
{
  "layout": "chart",
  "content": {
    "title": "頁面標題",
    "chart": {
      "type": "column",
      "data": {
        "categories": ["類別1", "類別2", "類別3"],
        "series": [
          {
            "name": "系列1",
            "values": [10, 20, 30]
          },
          {
            "name": "系列2",
            "values": [15, 25, 35]
          }
        ]
      }
    }
  }
}
```

**支持的圖表類型**:
- `column`: 柱狀圖(垂直)
- `bar`: 條形圖(水平)
- `line`: 折線圖
- `pie`: 圓餅圖

**注意**: 圓餅圖通常只使用一個系列的數據。

## 完整示例

查看 `example_config.json` 文件,其中包含了所有佈局類型的完整示例。

## 項目結構

```
ppt_assistant/
├── ppt_assistant/        # 核心模塊
│   ├── __init__.py       # 模塊初始化
│   ├── __main__.py       # 模塊主入口
│   └── core.py           # 核心功能實現
├── ppt_assistant.py      # 主腳本入口點
├── README.md             # 使用文檔
├── requirements.txt      # 依賴列表
├── .gitignore            # Git 忽略文件
├── examples/             # 示例文件
│   ├── example_config.json   # 示例配置文件
│   ├── sample_image.png      # 示例圖片
│   └── output_presentation.pptx  # 生成的示例簡報 (16:9)
├── docs/                 # 文檔
│   ├── FORMATTING_GUIDE.md   # 格式化指南
│   ├── json_format.md        # JSON 格式詳細說明
│   ├── LAYOUT_SPECS.md       # 佈局規範
│   ├── PROJECT_SUMMARY.md    # 項目總結
│   └── QUICKSTART.md         # 快速開始
└── tests/                # 測試文件
    └── __init__.py       # 測試模塊初始化
```

## 進階使用


### 批量生成簡報

您可以創建多個 JSON 配置文件,然後使用腳本批量生成:

```bash
for config in *.json; do
    python3 ppt_assistant.py "$config"
done
```

### 程式化調用

您也可以在自己的 Python 程序中導入並使用:

```python
from ppt_assistant import PPTAssistant

# 創建助手實例
assistant = PPTAssistant('config.json')

# 生成簡報
assistant.create_presentation()
```

## 技術規格

- **簡報比例**: 16:9 (寬屏)
- **簡報尺寸**: 13.333 x 7.5 英寸 (PowerPoint 標準寬屏尺寸)
- **支持格式**: .pptx (PowerPoint 2007+)
- **圖片格式**: PNG, JPG, JPEG, GIF, BMP 等
- **編碼**: UTF-8

## 注意事項

1. **圖片路徑**: 確保 JSON 中指定的圖片路徑正確且文件存在
2. **數據格式**: 表格和圖表的數據必須格式正確,否則可能導致生成失敗
4. **文件編碼**: JSON 文件必須使用 UTF-8 編碼,以支持中文等字符
5. **輸出路徑**: 確保輸出路徑的目錄存在且有寫入權限
6. **16:9 比例**: 所有生成的簡報都會自動設置為 16:9 比例,無需手動調整

## 常見問題


### Q: 圖片沒有顯示?

A: 檢查圖片路徑是否正確,建議使用絕對路徑或相對於執行目錄的路徑。

### Q: 圖表顯示不正確?

A: 確保數據格式正確,categories 和 values 的數量要匹配。

### Q: 支持哪些圖片格式?

A: 支持常見的圖片格式,包括 PNG、JPG、JPEG、GIF、BMP 等。

### Q: 可以改變簡報比例嗎?

A: 當前版本固定為 16:9 比例。如需其他比例,可以修改代碼中的 `slide_width` 和 `slide_height` 參數。

## 技術支持

如有問題或建議,請查看代碼中的註釋或聯繫開發者。

## 許可證

本工具基於 python-pptx 庫開發,遵循相應的開源許可證。
