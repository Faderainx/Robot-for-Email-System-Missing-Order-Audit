"""主程序入口"""
import sys
import os

# 确保项目根目录在 sys.path 中
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)


def main():
    from gui import main as gui_main
    gui_main()


if __name__ == "__main__":
    main()
