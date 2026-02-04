#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
DOCX to HTML Converter - PyQt 버전 진입점
"""

import sys
import os

# 현재 디렉토리를 path에 추가
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gui_pyqt import main

if __name__ == "__main__":
    main()
