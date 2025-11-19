"""
Configuration constants for PPT Assistant
All hardcoded values are extracted here for easy customization

Units:
- Dimensions: inches
- Font sizes: points
- Colors: RGB tuples (0-255)
- Spacing: inches or points as noted
"""

from typing import Dict, Tuple
from pptx.dml.color import RGBColor


class PPTConfig:
    """
    Configuration class for PPT Assistant styling and layout constants.

    This class centralizes all presentation styling parameters to eliminate
    magic numbers from the codebase and enable easy customization.

    All dimensions are in inches unless otherwise specified.
    Colors are RGB tuples with values 0-255.
    """

    # Slide dimensions (16:9 ratio)
    SLIDE_WIDTH: float = 13.333
    SLIDE_HEIGHT: float = 7.5

    # Layout margins and spacing
    MARGIN: float = 0.5
    TITLE_HEIGHT: float = 1.0
    CONTENT_TOP_OFFSET: float = 0.5  # Space between title and content

    # Slide layout indices
    TITLE_LAYOUT_INDEX: int = 0
    CONTENT_LAYOUT_INDEX: int = 6  # Fallback to 5 if not available
    CONTENT_LAYOUT_FALLBACK_INDEX: int = 5

    # Font sizes (in points)
    TITLE_FONT_SIZE: int = 54
    SUBTITLE_FONT_SIZE: int = 32
    SLIDE_TITLE_FONT_SIZE: int = 40
    BODY_FONT_SIZE: int = 24
    COLUMN_FONT_SIZE: int = 20
    CAPTION_FONT_SIZE: int = 18
    TABLE_HEADER_FONT_SIZE: int = 20
    TABLE_DATA_FONT_SIZE: int = 18

    # Spacing (in points)
    PARAGRAPH_SPACE_BEFORE: int = 8
    PARAGRAPH_SPACE_AFTER: int = 8
    COLUMN_SPACE_AFTER: int = 12

    # Layout-specific dimensions
    TITLE_TOP_POSITION: float = 2.5
    TITLE_HEIGHT_SIZE: float = 1.5
    SUBTITLE_TOP_POSITION: float = 4.5
    SUBTITLE_HEIGHT_SIZE: float = 1.0
    SLIDE_TITLE_TOP: float = 0.3
    SLIDE_TITLE_HEIGHT: float = 0.8

    # Two-column layout
    COLUMN_GAP: float = 0.333

    # Image layout
    IMAGE_WIDTH_RATIO: float = 0.8  # 80% of content width
    IMAGE_TOP_OFFSET: float = 0.2
    CAPTION_TOP_FROM_BOTTOM: float = 1.0
    CAPTION_HEIGHT: float = 0.6

    # Table layout
    TABLE_WIDTH_RATIO: float = 0.95  # 95% of content width
    TABLE_HEIGHT_RATIO: float = 0.9  # 90% of content height

    # Chart layout
    CHART_WIDTH_RATIO: float = 0.9  # 90% of content width
    CHART_HEIGHT_RATIO: float = 0.9  # 90% of content height
    CHART_LEGEND_POSITION: int = 2  # Right side

    # Colors (RGB tuples)
    TABLE_HEADER_BG_COLOR: Tuple[int, int, int] = (68, 114, 196)  # Blue
    TABLE_HEADER_TEXT_COLOR: Tuple[int, int, int] = (255, 255, 255)  # White
    TABLE_ALT_ROW_BG_COLOR: Tuple[int, int, int] = (217, 226, 243)  # Light blue

    # Chart type mappings
    CHART_TYPE_MAPPING: Dict[str, str] = {
        'bar': 'BAR_CLUSTERED',
        'column': 'COLUMN_CLUSTERED',
        'line': 'LINE',
        'pie': 'PIE'
    }

    # Slide dimensions (16:9 ratio)
    SLIDE_WIDTH = 13.333
    SLIDE_HEIGHT = 7.5

    # Layout margins and spacing
    MARGIN = 0.5
    TITLE_HEIGHT = 1.0
    CONTENT_TOP_OFFSET = 0.5  # Space between title and content

    def __init__(self):
        """Initialize with default values and validate configuration."""
        self._validate_config()

    def _validate_config(self):
        """Validate that configuration values are reasonable."""
        if self.SLIDE_WIDTH <= 2 * self.MARGIN:
            raise ValueError(f"Slide width ({self.SLIDE_WIDTH}) must be greater than 2 * margin ({2 * self.MARGIN})")

        if self.SLIDE_HEIGHT <= self.TITLE_HEIGHT + self.CONTENT_TOP_OFFSET + self.MARGIN:
            raise ValueError("Slide height too small for content layout")

        if not (0 < self.IMAGE_WIDTH_RATIO <= 1):
            raise ValueError("Image width ratio must be between 0 and 1")

        if not (0 < self.TABLE_WIDTH_RATIO <= 1):
            raise ValueError("Table width ratio must be between 0 and 1")

        if not (0 < self.TABLE_HEIGHT_RATIO <= 1):
            raise ValueError("Table height ratio must be between 0 and 1")

        if not (0 < self.CHART_WIDTH_RATIO <= 1):
            raise ValueError("Chart width ratio must be between 0 and 1")

        if not (0 < self.CHART_HEIGHT_RATIO <= 1):
            raise ValueError("Chart height ratio must be between 0 and 1")

        # Validate color values
        for color_name, color_value in [
            ("TABLE_HEADER_BG_COLOR", self.TABLE_HEADER_BG_COLOR),
            ("TABLE_HEADER_TEXT_COLOR", self.TABLE_HEADER_TEXT_COLOR),
            ("TABLE_ALT_ROW_BG_COLOR", self.TABLE_ALT_ROW_BG_COLOR)
        ]:
            if not all(0 <= c <= 255 for c in color_value):
                raise ValueError(f"{color_name} values must be between 0 and 255")

    # Calculated dimensions
    @property
    def content_width(self) -> float:
        """Calculate content width based on slide width and margins."""
        return self.SLIDE_WIDTH - (2 * self.MARGIN)

    @property
    def content_top(self) -> float:
        """Calculate top position of content area."""
        return self.TITLE_HEIGHT + self.CONTENT_TOP_OFFSET

    @property
    def content_height(self) -> float:
        """Calculate height of content area."""
        return self.SLIDE_HEIGHT - self.content_top - self.MARGIN

    @classmethod
    def from_json(cls, config_path: str) -> 'PPTConfig':
        """
        Load configuration from a JSON file.

        Args:
            config_path: Path to JSON configuration file

        Returns:
            PPTConfig instance with loaded values

        Raises:
            FileNotFoundError: If config file doesn't exist
            json.JSONDecodeError: If JSON is invalid
            ValueError: If config values are invalid
        """
        import json

        with open(config_path, 'r', encoding='utf-8') as f:
            config_data = json.load(f)

        # Create instance and override defaults
        instance = cls()

        # Override attributes from JSON
        for key, value in config_data.items():
            if hasattr(instance, key):
                setattr(instance, key, value)
            else:
                raise ValueError(f"Unknown configuration key: {key}")

        # Re-validate with new values
        instance._validate_config()

        return instance

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