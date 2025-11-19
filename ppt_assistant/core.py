import json
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, PP_PARAGRAPH_ALIGNMENT
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.dml.color import RGBColor


class PPTAssistant:
    """PPT 助手類"""

    # 16:9 簡報尺寸常量
    SLIDE_WIDTH = 13.333
    SLIDE_HEIGHT = 7.5
    MARGIN = 0.5
    CONTENT_WIDTH = SLIDE_WIDTH - (2 * MARGIN)  # 12.333
    TITLE_HEIGHT = 1.0
    CONTENT_TOP = TITLE_HEIGHT + 0.5  # 1.5
    CONTENT_HEIGHT = SLIDE_HEIGHT - CONTENT_TOP - MARGIN  # 6.0

    def __init__(self, json_file):
        """
        初始化 PPT 助手

        Args:
            json_file: JSON 配置文件路徑
        """
        with open(json_file, 'r', encoding='utf-8') as f:
            self.config = json.load(f)

        self.template_path = self.config.get('template')
        self.output_path = self.config.get('output')
        self.slides_data = self.config.get('slides', [])

        # 載入模板或創建新簡報
        if self.template_path and os.path.exists(self.template_path):
            self.prs = Presentation(self.template_path)
        else:
            self.prs = Presentation()

        # 設置為 16:9 比例 (PowerPoint 標準寬屏尺寸)
        self.prs.slide_width = Inches(self.SLIDE_WIDTH)
        self.prs.slide_height = Inches(self.SLIDE_HEIGHT)

    def create_presentation(self):
        """創建完整的簡報"""
        for slide_data in self.slides_data:
            layout_type = slide_data.get('layout')
            content = slide_data.get('content', {})

            if layout_type == 'title':
                self._add_title_slide(content)
            elif layout_type == 'content':
                self._add_content_slide(content)
            elif layout_type == 'two_column':
                self._add_two_column_slide(content)
            elif layout_type == 'image':
                self._add_image_slide(content)
            elif layout_type == 'table':
                self._add_table_slide(content)
            elif layout_type == 'chart':
                self._add_chart_slide(content)
            else:
                print(f"警告: 不支援的佈局類型 '{layout_type}'")

        # 保存簡報
        self.prs.save(self.output_path)
        print(f"簡報已成功生成: {self.output_path}")

    def _add_title_slide(self, content):
        """添加標題頁"""
        slide_layout = self.prs.slide_layouts[0]
        slide = self.prs.slides.add_slide(slide_layout)

        # 設置標題
        if slide.shapes.title:
            title = slide.shapes.title
            title.text = content.get('title', '')
            title.left = Inches(self.MARGIN)
            title.top = Inches(2.5)
            title.width = Inches(self.CONTENT_WIDTH)
            title.height = Inches(1.5)
            tf = title.text_frame
            tf.paragraphs[0].alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
            tf.paragraphs[0].font.size = Pt(54)
            tf.paragraphs[0].font.bold = True

        # 設置副標題
        if len(slide.placeholders) > 1:
            subtitle = slide.placeholders[1]
            subtitle.text = content.get('subtitle', '')
            subtitle.left = Inches(self.MARGIN)
            subtitle.top = Inches(4.5)
            subtitle.width = Inches(self.CONTENT_WIDTH)
            subtitle.height = Inches(1.0)
            tf = subtitle.text_frame
            tf.paragraphs[0].alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
            tf.paragraphs[0].font.size = Pt(32)

    def _add_content_slide(self, content):
        """添加內容頁"""
        slide_layout = self.prs.slide_layouts[6] if len(self.prs.slide_layouts) > 6 else self.prs.slide_layouts[5]
        slide = self.prs.slides.add_slide(slide_layout)

        # 添加標題
        title_shape = slide.shapes.add_textbox(
            Inches(self.MARGIN),
            Inches(0.3),
            Inches(self.CONTENT_WIDTH),
            Inches(0.8)
        )
        tf = title_shape.text_frame
        p = tf.paragraphs[0]
        p.text = content.get('title', '')
        p.font.size = Pt(40)
        p.font.bold = True
        p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT

        # 添加內容
        body_shape = slide.shapes.add_textbox(
            Inches(self.MARGIN),
            Inches(self.CONTENT_TOP),
            Inches(self.CONTENT_WIDTH),
            Inches(self.CONTENT_HEIGHT)
        )
        tf = body_shape.text_frame
        tf.word_wrap = True

        body_text = content.get('body', '')
        if isinstance(body_text, list):
            for idx, item in enumerate(body_text):
                if idx == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                p.text = str(item)
                p.level = 0
                p.font.size = Pt(24)
                p.space_before = Pt(8)
                p.space_after = Pt(8)
                p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT
        else:
            p = tf.paragraphs[0]
            p.text = str(body_text)
            p.font.size = Pt(24)
            p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT

    def _add_two_column_slide(self, content):
        """添加雙欄佈局頁"""
        slide_layout = self.prs.slide_layouts[6] if len(self.prs.slide_layouts) > 6 else self.prs.slide_layouts[5]
        slide = self.prs.slides.add_slide(slide_layout)

        # 添加標題
        title_shape = slide.shapes.add_textbox(
            Inches(self.MARGIN),
            Inches(0.3),
            Inches(self.CONTENT_WIDTH),
            Inches(0.8)
        )
        tf = title_shape.text_frame
        p = tf.paragraphs[0]
        p.text = content.get('title', '')
        p.font.size = Pt(40)
        p.font.bold = True
        p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT

        # 計算雙欄尺寸
        column_gap = 0.333
        column_width = (self.CONTENT_WIDTH - column_gap) / 2  # 6.0 inches each

        # 左欄
        left_box = slide.shapes.add_textbox(
            Inches(self.MARGIN),
            Inches(self.CONTENT_TOP),
            Inches(column_width),
            Inches(self.CONTENT_HEIGHT)
        )
        left_tf = left_box.text_frame
        left_tf.word_wrap = True
        left_p = left_tf.paragraphs[0]
        left_p.text = str(content.get('left', ''))
        left_p.font.size = Pt(20)
        left_p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT
        left_p.space_after = Pt(12)

        # 右欄
        right_left_pos = self.MARGIN + column_width + column_gap
        right_box = slide.shapes.add_textbox(
            Inches(right_left_pos),
            Inches(self.CONTENT_TOP),
            Inches(column_width),
            Inches(self.CONTENT_HEIGHT)
        )
        right_tf = right_box.text_frame
        right_tf.word_wrap = True
        right_p = right_tf.paragraphs[0]
        right_p.text = str(content.get('right', ''))
        right_p.font.size = Pt(20)
        right_p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT
        right_p.space_after = Pt(12)

    def _add_image_slide(self, content):
        """添加圖片頁"""
        slide_layout = self.prs.slide_layouts[6] if len(self.prs.slide_layouts) > 6 else self.prs.slide_layouts[5]
        slide = self.prs.slides.add_slide(slide_layout)

        # 添加標題
        title_shape = slide.shapes.add_textbox(
            Inches(self.MARGIN),
            Inches(0.3),
            Inches(self.CONTENT_WIDTH),
            Inches(0.8)
        )
        tf = title_shape.text_frame
        p = tf.paragraphs[0]
        p.text = content.get('title', '')
        p.font.size = Pt(40)
        p.font.bold = True
        p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT

        # 添加圖片
        image_path = content.get('image', '')
        if image_path and os.path.exists(image_path):
            # 圖片居中,使用大部分內容區域
            img_width = self.CONTENT_WIDTH * 0.8  # 使用 80% 的內容寬度
            img_left = self.MARGIN + (self.CONTENT_WIDTH - img_width) / 2
            img_top = self.CONTENT_TOP + 0.2
            slide.shapes.add_picture(
                image_path,
                Inches(img_left),
                Inches(img_top),
                width=Inches(img_width)
            )

        # 添加圖片說明(可選)
        caption = content.get('caption', '')
        if caption:
            caption_shape = slide.shapes.add_textbox(
                Inches(self.MARGIN),
                Inches(self.SLIDE_HEIGHT - 1.0),
                Inches(self.CONTENT_WIDTH),
                Inches(0.6)
            )
            caption_tf = caption_shape.text_frame
            caption_p = caption_tf.paragraphs[0]
            caption_p.text = caption
            caption_p.font.size = Pt(18)
            caption_p.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
            caption_p.font.italic = True

    def _add_table_slide(self, content):
        """添加表格頁"""
        slide_layout = self.prs.slide_layouts[6] if len(self.prs.slide_layouts) > 6 else self.prs.slide_layouts[5]
        slide = self.prs.slides.add_slide(slide_layout)

        # 添加標題
        title_shape = slide.shapes.add_textbox(
            Inches(self.MARGIN),
            Inches(0.3),
            Inches(self.CONTENT_WIDTH),
            Inches(0.8)
        )
        tf = title_shape.text_frame
        p = tf.paragraphs[0]
        p.text = content.get('title', '')
        p.font.size = Pt(40)
        p.font.bold = True
        p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT

        # 獲取表格數據
        table_data = content.get('table', {})
        headers = table_data.get('headers', [])
        rows = table_data.get('rows', [])

        if headers and rows:
            # 計算表格尺寸
            rows_count = len(rows) + 1  # +1 for header
            cols_count = len(headers)

            # 添加表格 - 使用完整內容區域
            table_width = self.CONTENT_WIDTH * 0.95
            table_left = self.MARGIN + (self.CONTENT_WIDTH - table_width) / 2

            table_shape = slide.shapes.add_table(
                rows_count,
                cols_count,
                Inches(table_left),
                Inches(self.CONTENT_TOP),
                Inches(table_width),
                Inches(self.CONTENT_HEIGHT * 0.9)
            )
            table = table_shape.table

            # 填充表頭
            for col_idx, header in enumerate(headers):
                cell = table.cell(0, col_idx)
                cell.text = str(header)
                # 格式化表頭
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.font.bold = True
                    paragraph.font.size = Pt(20)
                    paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                # 設置表頭背景色
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(68, 114, 196)
                # 設置文字顏色為白色
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.font.color.rgb = RGBColor(255, 255, 255)

            # 填充數據行
            for row_idx, row_data in enumerate(rows):
                for col_idx, cell_data in enumerate(row_data):
                    cell = table.cell(row_idx + 1, col_idx)
                    cell.text = str(cell_data)
                    # 格式化數據單元格
                    for paragraph in cell.text_frame.paragraphs:
                        paragraph.font.size = Pt(18)
                        paragraph.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                    # 交替行背景色
                    if row_idx % 2 == 0:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(217, 226, 243)

    def _add_chart_slide(self, content):
        """添加圖表頁"""
        slide_layout = self.prs.slide_layouts[6] if len(self.prs.slide_layouts) > 6 else self.prs.slide_layouts[5]
        slide = self.prs.slides.add_slide(slide_layout)

        # 添加標題
        title_shape = slide.shapes.add_textbox(
            Inches(self.MARGIN),
            Inches(0.3),
            Inches(self.CONTENT_WIDTH),
            Inches(0.8)
        )
        tf = title_shape.text_frame
        p = tf.paragraphs[0]
        p.text = content.get('title', '')
        p.font.size = Pt(40)
        p.font.bold = True
        p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT

        # 獲取圖表數據
        chart_config = content.get('chart', {})
        chart_type = chart_config.get('type', 'bar')
        chart_data_config = chart_config.get('data', {})

        categories = chart_data_config.get('categories', [])
        series_list = chart_data_config.get('series', [])

        if categories and series_list:
            # 創建圖表數據
            chart_data = CategoryChartData()
            chart_data.categories = categories

            for series in series_list:
                series_name = series.get('name', '')
                series_values = series.get('values', [])
                chart_data.add_series(series_name, series_values)

            # 確定圖表類型
            chart_type_map = {
                'bar': XL_CHART_TYPE.BAR_CLUSTERED,
                'column': XL_CHART_TYPE.COLUMN_CLUSTERED,
                'line': XL_CHART_TYPE.LINE,
                'pie': XL_CHART_TYPE.PIE
            }
            xl_chart_type = chart_type_map.get(chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)

            # 添加圖表 - 使用完整內容區域
            chart_width = self.CONTENT_WIDTH * 0.9
            chart_left = self.MARGIN + (self.CONTENT_WIDTH - chart_width) / 2

            chart_shape = slide.shapes.add_chart(
                xl_chart_type,
                Inches(chart_left),
                Inches(self.CONTENT_TOP),
                Inches(chart_width),
                Inches(self.CONTENT_HEIGHT * 0.9),
                chart_data
            )
            chart = chart_shape.chart

            # 設置圖表樣式
            chart.has_legend = True
            chart.legend.position = 2  # Right
            chart.legend.include_in_layout = False