#!/usr/bin/env python3
"""
바코드 라벨 생성기 메인 실행 파일
"""

import sys
from pathlib import Path

# 현재 디렉토리를 Python 경로에 추가
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

try:

    from PyQt6.QtGui import QPalette, QColor
    from PyQt6.QtCore import Qt
    from PyQt6.QtWidgets import QApplication
    from src.ui.main_window import MainWindow

    def force_light_theme(app: QApplication):
        # 1) 시스템 스타일 대신 Fusion 고정
        app.setStyle("Fusion")

        # 2) 밝은 팔레트 직접 지정 (필요한 역할만 골라 설정)
        pal = QPalette()

        pal.setColor(QPalette.ColorRole.Window, QColor("#ffffff"))
        pal.setColor(QPalette.ColorRole.Base, QColor("#ffffff"))           # 입력창/표 바탕
        pal.setColor(QPalette.ColorRole.AlternateBase, QColor("#f6f6f6"))
        pal.setColor(QPalette.ColorRole.Text, QColor("#000000"))           # 일반 텍스트
        pal.setColor(QPalette.ColorRole.WindowText, QColor("#000000"))     # 라벨 등
        pal.setColor(QPalette.ColorRole.Button, QColor("#ffffff"))
        pal.setColor(QPalette.ColorRole.ButtonText, QColor("#000000"))
        pal.setColor(QPalette.ColorRole.Highlight, QColor("#0078d4"))
        pal.setColor(QPalette.ColorRole.HighlightedText, QColor("#ffffff"))
        pal.setColor(QPalette.ColorRole.ToolTipBase, QColor("#ffffdc"))
        pal.setColor(QPalette.ColorRole.ToolTipText, QColor("#000000"))
        pal.setColor(QPalette.ColorRole.PlaceholderText, QColor("#666666"))

        app.setPalette(pal)


    def main():
        """메인 함수"""
        # QApplication 생성
        app = QApplication(sys.argv)
        force_light_theme(app)
        app.setApplicationName("바코드 라벨 생성기")
        app.setApplicationVersion("1.0")
        app.setOrganizationName("Barcode Label Generator")

        # 고해상도 디스플레이 지원 (PyQt6에서는 기본적으로 활성화됨)
        # 필요한 경우에만 설정
        try:
            app.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
            app.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)
        except AttributeError:
            pass  # 일부 PyQt6 버전에서는 필요 없을 수 있음

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