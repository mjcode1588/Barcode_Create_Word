#!/usr/bin/env python3
"""
바코드 라벨 생성기 실행 스크립트
"""

import sys
import os
from pathlib import Path

# 현재 디렉토리를 Python 경로에 추가
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

# src 디렉토리를 Python 경로에 추가
src_dir = current_dir / "src"
sys.path.insert(0, str(src_dir))

try:
    from PyQt6.QtWidgets import QApplication
    from PyQt6.QtCore import Qt
    from ui.main_window import MainWindow
    
    def main():
        """메인 함수"""
        # QApplication 생성
        app = QApplication(sys.argv)
        app.setApplicationName("바코드 라벨 생성기")
        app.setApplicationVersion("1.0")
        app.setOrganizationName("Barcode Label Generator")
        
        # 고해상도 디스플레이 지원 (PyQt6에서는 기본적으로 활성화됨)
        # 필요한 경우에만 설정
        try:
            app.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
            app.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)
        except:
            pass  # PyQt6에서는 기본적으로 지원되므로 무시
        
        # 메인 윈도우 생성 및 표시
        main_window = MainWindow()
        main_window.show()
        
        # 이벤트 루프 시작
        sys.exit(app.exec())
    
    if __name__ == "__main__":
        main()
        
except ImportError as e:
    print(f"Import 오류: {e}")
    print("PyQt6가 설치되어 있는지 확인하세요.")
    print("uv sync 명령을 실행하여 의존성을 설치하세요.")
    sys.exit(1)
except Exception as e:
    print(f"오류 발생: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
