# PPT 佈局尺寸規範 (16:9)

## 簡報基本尺寸

- **簡報寬度**: 13.333 inches
- **簡報高度**: 7.5 inches
- **寬高比**: 1.7778:1 (16:9)
- **邊距**: 0.5 inches
- **內容區域寬度**: 12.333 inches (簡報寬度 - 2 × 邊距)
- **內容區域高度**: 6.0 inches (簡報高度 - 標題區域 - 底部邊距)

## 尺寸常量定義

```python
SLIDE_WIDTH = 13.333      # 簡報寬度
SLIDE_HEIGHT = 7.5        # 簡報高度
MARGIN = 0.5              # 邊距
CONTENT_WIDTH = 12.333    # 內容區域寬度
TITLE_HEIGHT = 1.0        # 標題區域高度
CONTENT_TOP = 1.5         # 內容起始位置
CONTENT_HEIGHT = 6.0      # 內容區域高度
```

## 各類型投影片佈局規範

### 1. 標題頁 (Title Slide)

**標題文字框**:
- 位置: (0.5, 2.5)
- 尺寸: 12.333 × 1.5 inches
- 字體: 54pt 粗體
- 對齊: 居中

**副標題文字框**:
- 位置: (0.5, 4.5)
- 尺寸: 12.333 × 1.0 inches
- 字體: 32pt
- 對齊: 居中

### 2. 內容頁 (Content Slide)

**標題文字框**:
- 位置: (0.5, 0.3)
- 尺寸: 12.333 × 0.8 inches
- 字體: 40pt 粗體
- 對齊: 左對齊

**內容文字框**:
- 位置: (0.5, 1.5)
- 尺寸: 12.333 × 6.0 inches
- 字體: 24pt
- 對齊: 左對齊
- 列表項間距: 8pt

### 3. 雙欄佈局 (Two Column Slide)

**標題文字框**:
- 位置: (0.5, 0.3)
- 尺寸: 12.333 × 0.8 inches
- 字體: 40pt 粗體
- 對齊: 左對齊

**左欄文字框**:
- 位置: (0.5, 1.5)
- 尺寸: 6.0 × 6.0 inches
- 字體: 20pt
- 對齊: 左對齊

**右欄文字框**:
- 位置: (6.833, 1.5)
- 尺寸: 6.0 × 6.0 inches
- 字體: 20pt
- 對齊: 左對齊

**欄間距**: 0.333 inches

**計算公式**:
```
column_width = (CONTENT_WIDTH - column_gap) / 2
             = (12.333 - 0.333) / 2
             = 6.0 inches

right_column_left = MARGIN + column_width + column_gap
                  = 0.5 + 6.0 + 0.333
                  = 6.833 inches
```

### 4. 圖片頁 (Image Slide)

**標題文字框**:
- 位置: (0.5, 0.3)
- 尺寸: 12.333 × 0.8 inches
- 字體: 40pt 粗體
- 對齊: 左對齊

**圖片**:
- 寬度: 9.866 inches (內容寬度的 80%)
- 位置: 水平居中
- 左邊距: 1.733 inches (計算得出)
- 上邊距: 1.7 inches

**圖片說明文字框**:
- 位置: (0.5, 6.5)
- 尺寸: 12.333 × 0.6 inches
- 字體: 18pt 斜體
- 對齊: 居中

### 5. 表格頁 (Table Slide)

**標題文字框**:
- 位置: (0.5, 0.3)
- 尺寸: 12.333 × 0.8 inches
- 字體: 40pt 粗體
- 對齊: 左對齊

**表格**:
- 寬度: 11.716 inches (內容寬度的 95%)
- 位置: 水平居中
- 左邊距: 0.809 inches (計算得出)
- 上邊距: 1.5 inches
- 高度: 5.4 inches (內容高度的 90%)

**表格樣式**:
- 表頭字體: 20pt 粗體,白色
- 表頭背景: RGB(68, 114, 196) 藍色
- 數據字體: 18pt,居中
- 交替行背景: RGB(217, 226, 243) 淺藍色

### 6. 圖表頁 (Chart Slide)

**標題文字框**:
- 位置: (0.5, 0.3)
- 尺寸: 12.333 × 0.8 inches
- 字體: 40pt 粗體
- 對齊: 左對齊

**圖表**:
- 寬度: 11.1 inches (內容寬度的 90%)
- 位置: 水平居中
- 左邊距: 1.117 inches (計算得出)
- 上邊距: 1.5 inches
- 高度: 5.4 inches (內容高度的 90%)

**圖表設置**:
- 圖例: 顯示在右側
- 圖例不包含在佈局中

## 尺寸計算範例

### 水平居中計算

對於需要居中的元素:

```python
element_width = CONTENT_WIDTH * percentage  # 例如 0.8 或 0.9
left_position = MARGIN + (CONTENT_WIDTH - element_width) / 2
```

範例 (圖片 80% 寬度):
```python
img_width = 12.333 * 0.8 = 9.866 inches
img_left = 0.5 + (12.333 - 9.866) / 2 = 1.733 inches
```

### 雙欄分配計算

```python
column_gap = 0.333 inches
column_width = (CONTENT_WIDTH - column_gap) / 2
left_column_left = MARGIN
right_column_left = MARGIN + column_width + column_gap
```

## 與 4:3 比例的對比

| 項目 | 4:3 比例 | 16:9 比例 | 差異 |
|-----|---------|-----------|------|
| 簡報寬度 | 10.0" | 13.333" | +33.3% |
| 簡報高度 | 7.5" | 7.5" | 相同 |
| 內容寬度 | 9.0" | 12.333" | +37.0% |
| 寬高比 | 1.333:1 | 1.778:1 | +33.3% |

**關鍵差異**: 16:9 比例提供了更寬的水平空間,特別適合:
- 雙欄佈局
- 寬幅圖表
- 橫向圖片
- 表格展示

## 設計建議

1. **充分利用寬度**: 16:9 提供更多水平空間,可以展示更多內容
2. **避免過度擁擠**: 雖然空間更大,但仍要保持適當留白
3. **保持一致性**: 所有投影片使用相同的邊距和對齊方式
4. **文字大小**: 由於投影片更寬,可以適當增加字體大小以保持可讀性
5. **視覺平衡**: 注意左右平衡,避免內容偏向一側

## 驗證方法

使用以下代碼驗證生成的 PPT 尺寸:

```python
from pptx import Presentation

prs = Presentation('output.pptx')
print(f'寬度: {prs.slide_width.inches:.3f} inches')
print(f'高度: {prs.slide_height.inches:.3f} inches')
print(f'比例: {prs.slide_width.inches / prs.slide_height.inches:.4f}:1')

# 檢查投影片元素
for slide in prs.slides:
    for shape in slide.shapes:
        if hasattr(shape, 'left'):
            print(f'位置: ({shape.left.inches:.2f}, {shape.top.inches:.2f})')
            print(f'尺寸: {shape.width.inches:.2f} x {shape.height.inches:.2f}')
```

## 總結

所有元素尺寸現已完全適配 16:9 比例,確保:
- ✓ 文字框寬度使用完整的 12.333 inches 內容區域
- ✓ 雙欄佈局平均分配空間 (各 6.0 inches)
- ✓ 圖片、表格、圖表使用合適的比例 (80-95%)
- ✓ 所有元素正確對齊,不會出現 4:3 比例的錯位問題
