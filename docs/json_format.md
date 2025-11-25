# PPT 助手 JSON 格式規範

## 整體結構

```json
{
  "output": "path/to/output.pptx",
  "footer": "Footer text (optional)",
  "slides": [
    {
      "layout": "title",
      "content": {...}
    },
    {
      "layout": "content",
      "content": {...}
    }
  ]
}
```

## 支援的佈局類型

### 1. title - 標題頁
```json
{
  "layout": "title",
  "content": {
    "title": "簡報標題",
    "subtitle": "副標題"
  }
}
```

### 2. content - 內容頁
```json
{
  "layout": "content",
  "content": {
    "title": "頁面標題",
    "body": "文字內容或項目列表"
  }
}
```

### 3. two_column - 雙欄佈局
```json
{
  "layout": "two_column",
  "content": {
    "title": "頁面標題",
    "left": "左欄內容",
    "right": "右欄內容"
  }
}
```

### 4. image - 圖片頁
```json
{
  "layout": "image",
  "content": {
    "title": "頁面標題",
    "image": "path/to/image.png",
    "caption": "圖片說明(可選)"
  }
}
```

### 5. table - 表格頁
```json
{
  "layout": "table",
  "content": {
    "title": "頁面標題",
    "table": {
      "headers": ["欄位1", "欄位2", "欄位3"],
      "rows": [
        ["數據1", "數據2", "數據3"],
        ["數據4", "數據5", "數據6"]
      ]
    }
  }
}
```

### 6. chart - 圖表頁
```json
{
  "layout": "chart",
  "content": {
    "title": "頁面標題",
    "chart": {
      "type": "bar",  // 支援: bar, column, line, pie
      "data": {
        "categories": ["類別1", "類別2", "類別3"],
        "series": [
          {
            "name": "系列1",
            "values": [10, 20, 30]
          }
        ]
      }
    }
  }
}
```

## 完整示例

```json
{
  "output": "output.pptx",
  "footer": "© 2024 Company Name",
  "slides": [
    {
      "layout": "title",
      "content": {
        "title": "2024 年度報告",
        "subtitle": "業務發展與成果展示"
      }
    },
    {
      "layout": "content",
      "content": {
        "title": "重點摘要",
        "body": "今年度我們達成了多項重要里程碑，包括業績成長、市場擴展和產品創新。"
      }
    },
    {
      "layout": "table",
      "content": {
        "title": "季度業績",
        "table": {
          "headers": ["季度", "營收", "成長率"],
          "rows": [
            ["Q1", "100萬", "10%"],
            ["Q2", "120萬", "20%"],
            ["Q3", "150萬", "25%"],
            ["Q4", "180萬", "20%"]
          ]
        }
      }
    }
  ]
}
```
