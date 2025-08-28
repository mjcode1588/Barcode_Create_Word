#!/usr/bin/env python3
"""
바코드 라벨 생성기 메인 실행 파일
"""

import sys
import os
from pathlib import Path

# 프로젝트 루트를 Python 경로에 추가
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt
from .ui.main_window import MainWindow

def main():
    """메인 함수"""
    # QApplication 생성
    app = QApplication(sys.argv)
    app.setApplicationName("바코드 라벨 생성기")
    app.setApplicationVersion("1.0")
    app.setOrganizationName("Barcode Label Generator")
    
    # 고해상도 디스플레이 지원
    app.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
    app.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)
    
    # 메인 윈도우 생성 및 표시
    main_window = MainWindow()
    main_window.show()
    
    # 이벤트 루프 시작
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
