"""
PPT 助手工具包
支持從 JSON 格式讀取內容,套用 .pptx 模板,並輸出包含文字、圖片、表格、圖表的 PPT 文件
完全適配 16:9 比例 (13.333 x 7.5 inches)
"""

from .core import PPTAssistant
from .config import PPTConfig, DEFAULT_CONFIG

__version__ = "1.0.0"
__all__ = ["PPTAssistant", "PPTConfig", "DEFAULT_CONFIG"]