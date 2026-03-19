# -*- coding: utf-8 -*-
"""
养老金融通报数据自动化处理系统
主程序入口

作者: Matrix Agent
创建日期: 2026-03-19
"""

import sys
import os

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gui.main_window import main


if __name__ == "__main__":
    main()
