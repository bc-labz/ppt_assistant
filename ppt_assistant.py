#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT 助手工具 - 主入口點
"""

import os
from ppt_assistant import PPTAssistant


def main():
    """主函數"""
    import sys

    if len(sys.argv) < 2:
        print("使用方法: python3 ppt_assistant.py <json_file>")
        print("範例: python3 ppt_assistant.py config.json")
        sys.exit(1)

    json_file = sys.argv[1]

    if not os.path.exists(json_file):
        print(f"錯誤: 找不到文件 '{json_file}'")
        sys.exit(1)

    try:
        assistant = PPTAssistant(json_file)
        assistant.create_presentation()
    except Exception as e:
        print(f"錯誤: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
