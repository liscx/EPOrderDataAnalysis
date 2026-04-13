import sys
import os

# 1. 确保项目根目录在 sys.path 中 (当前 GUI.py 就在根目录，直接取 dirname 即可)
root_dir = os.path.dirname(os.path.abspath(__file__))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

from src.gui.gui import ModernAnalysisGUI

if __name__ == "__main__":
    # 程序入口，启动 ModernAnalysisGUI 窗口
    # 注意：配置文件的创建逻辑已封合在 ModernAnalysisGUI 的 __init__ 方法中，
    # 当探测到根目录下无 config.yaml 时会自动根据当前日期生成填充默认值。
    app = ModernAnalysisGUI()
    app.mainloop()


