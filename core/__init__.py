# -*- coding: utf-8 -*-
"""
养老金融通报数据自动化处理系统
核心模块
"""

from .config import BRANCHES, BRANCH_COUNT, SINGLE_KEY_COLS, COLLECTIVE_KEY_COLS
from .processor import DataProcessor, processor

__all__ = [
    'BRANCHES',
    'BRANCH_COUNT',
    'SINGLE_KEY_COLS',
    'COLLECTIVE_KEY_COLS',
    'DataProcessor',
    'processor',
]
