"""Entry point for running ppt-assistant as a module."""

from .core import PPTAssistant
import sys
import os


def main():
    """主函數"""
    if len(sys.argv) < 2:
        print("使用方法: python -m ppt_assistant <json_file>")
        print("範例: python -m ppt_assistant examples/example_config.json")
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