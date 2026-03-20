import sys
import os

# 确保项目根目录在 sys.path 中
root_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

from src.gui.gui import ModernAnalysisGUI

if __name__ == "__main__":
    app = ModernAnalysisGUI()
    app.mainloop()
