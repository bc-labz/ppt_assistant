# 快速開始指南

## 5 分鐘快速上手

### 步驟 1: 安裝依賴

```bash
sudo pip3 install python-pptx
```

### 步驟 2: 準備配置文件

創建一個 JSON 文件 `my_presentation.json`:

```json
{
  "output": "my_first_presentation.pptx",
  "slides": [
    {
      "layout": "title",
      "content": {
        "title": "我的第一個簡報",
        "subtitle": "使用 PPT 助手工具創建"
      }
    },
    {
      "layout": "content",
      "content": {
        "title": "簡介",
        "body": [
          "這是一個自動生成的簡報",
          "使用 JSON 格式定義內容",
          "支持多種佈局和元素"
        ]
      }
    }
  ]
}
```

### 步驟 3: 生成簡報

```bash
python3 ppt_assistant.py my_presentation.json
```

### 步驟 4: 查看結果

生成的簡報文件為 `my_first_presentation.pptx`,可以用 PowerPoint 或其他兼容軟件打開。

## 運行示例

項目已包含完整的示例,直接運行:

```bash
python3 ppt_assistant.py example_config.json
```

這將生成一個包含所有佈局類型的示例簡報 `output_presentation.pptx`。

## 下一步

- 查看 `README.md` 了解詳細功能
- 查看 `json_format.md` 了解 JSON 格式規範
- 修改 `example_config.json` 嘗試不同的配置

## 常用命令

```bash
# 查看幫助
python3 ppt_assistant.py

# 使用配置文件生成簡報
python3 ppt_assistant.py config.json

# 查看生成的文件
ls -lh *.pptx
```

## 提示

1. JSON 文件必須使用 UTF-8 編碼
2. 圖片路徑建議使用相對路徑
3. 先用示例配置測試,確保工具正常運行
4. 逐步添加投影片,避免一次性配置過於複雜
