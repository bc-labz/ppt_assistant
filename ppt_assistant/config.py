"""
Configuration constants for PPT Assistant
All hardcoded values are extracted here for easy customization
"""

from pptx.dml.color import RGBColor


class PPTConfig:
    """Configuration class containing all PPT layout and styling constants"""

    # Slide dimensions (16:9 ratio)
    SLIDE_WIDTH = 13.333
    SLIDE_HEIGHT = 7.5

    # Layout margins and spacing
    MARGIN = 0.5
    TITLE_HEIGHT = 1.0
    CONTENT_TOP_OFFSET = 0.5  # Space between title and content

    # Calculated dimensions
    @property
    def content_width(self):
        return self.SLIDE_WIDTH - (2 * self.MARGIN)

    @property
    def content_top(self):
        return self.TITLE_HEIGHT + self.CONTENT_TOP_OFFSET

    @property
    def content_height(self):
        return self.SLIDE_HEIGHT - self.content_top - self.MARGIN

    # Slide layout indices
    TITLE_LAYOUT_INDEX = 0
    CONTENT_LAYOUT_INDEX = 6  # Fallback to 5 if not available
    CONTENT_LAYOUT_FALLBACK_INDEX = 5

    # Font sizes (in points)
    TITLE_FONT_SIZE = 54
    SUBTITLE_FONT_SIZE = 32
    SLIDE_TITLE_FONT_SIZE = 40
    BODY_FONT_SIZE = 24
    COLUMN_FONT_SIZE = 20
    CAPTION_FONT_SIZE = 18
    TABLE_HEADER_FONT_SIZE = 20
    TABLE_DATA_FONT_SIZE = 18

    # Spacing (in points)
    PARAGRAPH_SPACE_BEFORE = 8
    PARAGRAPH_SPACE_AFTER = 8
    COLUMN_SPACE_AFTER = 12

    # Layout-specific dimensions
    TITLE_TOP_POSITION = 2.5
    TITLE_HEIGHT_SIZE = 1.5
    SUBTITLE_TOP_POSITION = 4.5
    SUBTITLE_HEIGHT_SIZE = 1.0
    SLIDE_TITLE_TOP = 0.3
    SLIDE_TITLE_HEIGHT = 0.8

    # Two-column layout
    COLUMN_GAP = 0.333

    # Image layout
    IMAGE_WIDTH_RATIO = 0.8  # 80% of content width
    IMAGE_TOP_OFFSET = 0.2
    CAPTION_TOP_FROM_BOTTOM = 1.0
    CAPTION_HEIGHT = 0.6

    # Table layout
    TABLE_WIDTH_RATIO = 0.95  # 95% of content width
    TABLE_HEIGHT_RATIO = 0.9  # 90% of content height

    # Chart layout
    CHART_WIDTH_RATIO = 0.9  # 90% of content width
    CHART_HEIGHT_RATIO = 0.9  # 90% of content height
    CHART_LEGEND_POSITION = 2  # Right side

    # Colors (RGB tuples)
    TABLE_HEADER_BG_COLOR = (68, 114, 196)  # Blue
    TABLE_HEADER_TEXT_COLOR = (255, 255, 255)  # White
    TABLE_ALT_ROW_BG_COLOR = (217, 226, 243)  # Light blue

    # Chart type mappings
    CHART_TYPE_MAPPING = {
        'bar': 'BAR_CLUSTERED',
        'column': 'COLUMN_CLUSTERED',
        'line': 'LINE',
        'pie': 'PIE'
    }


# Default configuration instance
DEFAULT_CONFIG = PPTConfig()