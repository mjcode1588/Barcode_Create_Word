#!/usr/bin/env python3
"""
관리자 로그 다이얼로그 - 모든 애플리케이션 로그를 표시하고 관리
"""

from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QTextEdit, 
                               QPushButton, QComboBox, QLineEdit, QLabel, 
                               QGroupBox, QGridLayout, QFileDialog, QMessageBox,
                               QCheckBox, QSplitter, QListWidget, QListWidgetItem)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QFont, QTextCursor
from datetime import datetime
from typing import Optional

from src.services.log_service import logger, LogLevel, LogEntry


class AdminLogDialog(QDialog):
    """관리자 로그 다이얼로그"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("관리자 로그")
        self.setGeometry(200, 200, 1000, 700)
        self.setModal(False)  # 모달이 아닌 창으로 설정
        
        self.setup_ui()
        self.setup_connections()
        self.load_logs()
        
        # 자동 새로고침 타이머
        self.refresh_timer = QTimer()
        self.refresh_timer.timeout.connect(self.refresh_logs)
        self.auto_refresh = True
        self.refresh_timer.start(1000)  # 1초마다 새로고침
    
    def setup_ui(self):
        """UI 구성"""
        layout = QVBoxLayout()
        
        # 상단 컨트롤 패널
        control_group = QGroupBox("로그 제어")
        control_layout = QGridLayout()
        
        # 필터 컨트롤
        control_layout.addWidget(QLabel("레벨 필터:"), 0, 0)
        self.level_filter = QComboBox()
        self.level_filter.addItems(["전체", LogLevel.DEBUG, LogLevel.INFO, LogLevel.WARNING, LogLevel.ERROR, LogLevel.CRITICAL])
        control_layout.addWidget(self.level_filter, 0, 1)
        
        control_layout.addWidget(QLabel("모듈 필터:"), 0, 2)
        self.module_filter = QLineEdit()
        self.module_filter.setPlaceholderText("모듈명으로 검색...")
        control_layout.addWidget(self.module_filter, 0, 3)
        
        # 버튼들
        self.refresh_button = QPushButton("새로고침")
        control_layout.addWidget(self.refresh_button, 0, 4)
        
        self.auto_refresh_checkbox = QCheckBox("자동 새로고침")
        self.auto_refresh_checkbox.setChecked(True)
        control_layout.addWidget(self.auto_refresh_checkbox, 0, 5)
        
        # 두 번째 행
        self.clear_button = QPushButton("로그 삭제")
        control_layout.addWidget(self.clear_button, 1, 0)
        
        self.export_button = QPushButton("로그 내보내기")
        control_layout.addWidget(self.export_button, 1, 1)
        
        self.open_log_folder_button = QPushButton("로그 폴더 열기")
        control_layout.addWidget(self.open_log_folder_button, 1, 2)
        
        # 로그 통계
        self.log_count_label = QLabel("총 로그: 0개")
        control_layout.addWidget(self.log_count_label, 1, 3, 1, 3)
        
        control_group.setLayout(control_layout)
        layout.addWidget(control_group)
        
        # 메인 로그 영역 (분할 창)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # 왼쪽: 로그 목록
        log_list_group = QGroupBox("로그 목록")
        log_list_layout = QVBoxLayout()
        
        self.log_list = QListWidget()
        self.log_list.setMaximumWidth(300)
        log_list_layout.addWidget(self.log_list)
        
        log_list_group.setLayout(log_list_layout)
        splitter.addWidget(log_list_group)
        
        # 오른쪽: 로그 상세 내용
        log_detail_group = QGroupBox("로그 상세")
        log_detail_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        log_detail_layout.addWidget(self.log_text)
        
        log_detail_group.setLayout(log_detail_layout)
        splitter.addWidget(log_detail_group)
        
        splitter.setSizes([300, 700])
        layout.addWidget(splitter)
        
        # 하단 버튼
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        self.close_button = QPushButton("닫기")
        button_layout.addWidget(self.close_button)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def setup_connections(self):
        """시그널 연결"""
        self.level_filter.currentTextChanged.connect(self.apply_filters)
        self.module_filter.textChanged.connect(self.apply_filters)
        self.refresh_button.clicked.connect(self.refresh_logs)
        self.auto_refresh_checkbox.toggled.connect(self.toggle_auto_refresh)
        self.clear_button.clicked.connect(self.clear_logs)
        self.export_button.clicked.connect(self.export_logs)
        self.open_log_folder_button.clicked.connect(self.open_log_folder)
        self.close_button.clicked.connect(self.close)
        self.log_list.currentRowChanged.connect(self.show_log_detail)
        
        # 로그 서비스 시그널 연결
        logger.log_added.connect(self.on_new_log)
    
    def load_logs(self):
        """로그 로드"""
        self.apply_filters()
    
    def apply_filters(self):
        """필터 적용하여 로그 표시"""
        level_filter = self.level_filter.currentText()
        module_filter = self.module_filter.text().strip()
        
        # 필터 적용
        level_filter = None if level_filter == "전체" else level_filter
        module_filter = None if not module_filter else module_filter
        
        filtered_logs = logger.get_logs(level_filter, module_filter)
        
        # 로그 목록 업데이트
        self.log_list.clear()
        self.current_logs = filtered_logs
        
        for i, log in enumerate(filtered_logs):
            item = QListWidgetItem()
            item.setText(f"[{log.timestamp.strftime('%H:%M:%S')}] [{log.level}] {log.module}")
            
            # 레벨에 따른 색상 설정
            if log.level == LogLevel.ERROR or log.level == LogLevel.CRITICAL:
                item.setBackground(Qt.GlobalColor.red)
                item.setForeground(Qt.GlobalColor.white)
            elif log.level == LogLevel.WARNING:
                item.setBackground(Qt.GlobalColor.yellow)
            elif log.level == LogLevel.DEBUG:
                item.setForeground(Qt.GlobalColor.gray)
            
            self.log_list.addItem(item)
        
        # 통계 업데이트
        self.log_count_label.setText(f"총 로그: {len(filtered_logs)}개")
        
        # 마지막 로그로 스크롤
        if filtered_logs:
            self.log_list.setCurrentRow(len(filtered_logs) - 1)
    
    def show_log_detail(self, row: int):
        """선택된 로그의 상세 정보 표시"""
        if row >= 0 and row < len(self.current_logs):
            log = self.current_logs[row]
            
            detail_text = f"""시간: {log.timestamp.strftime('%Y-%m-%d %H:%M:%S')}
레벨: {log.level}
모듈: {log.module}
메시지: {log.message}

전체 로그:
{str(log)}"""
            
            self.log_text.setPlainText(detail_text)
    
    def refresh_logs(self):
        """로그 새로고침"""
        if self.auto_refresh:
            self.apply_filters()
    
    def toggle_auto_refresh(self, enabled: bool):
        """자동 새로고침 토글"""
        self.auto_refresh = enabled
        if enabled:
            self.refresh_timer.start(1000)
        else:
            self.refresh_timer.stop()
    
    def on_new_log(self, log_entry: LogEntry):
        """새 로그 엔트리 추가 시 호출"""
        if self.auto_refresh:
            # 현재 필터에 맞는지 확인
            level_filter = self.level_filter.currentText()
            module_filter = self.module_filter.text().strip()
            
            level_match = (level_filter == "전체" or log_entry.level == level_filter)
            module_match = (not module_filter or module_filter.lower() in log_entry.module.lower())
            
            if level_match and module_match:
                # 새 로그를 목록에 추가
                item = QListWidgetItem()
                item.setText(f"[{log_entry.timestamp.strftime('%H:%M:%S')}] [{log_entry.level}] {log_entry.module}")
                
                # 레벨에 따른 색상 설정
                if log_entry.level == LogLevel.ERROR or log_entry.level == LogLevel.CRITICAL:
                    item.setBackground(Qt.GlobalColor.red)
                    item.setForeground(Qt.GlobalColor.white)
                elif log_entry.level == LogLevel.WARNING:
                    item.setBackground(Qt.GlobalColor.yellow)
                elif log_entry.level == LogLevel.DEBUG:
                    item.setForeground(Qt.GlobalColor.gray)
                
                self.log_list.addItem(item)
                self.current_logs.append(log_entry)
                
                # 통계 업데이트
                self.log_count_label.setText(f"총 로그: {len(self.current_logs)}개")
                
                # 마지막 로그로 스크롤
                self.log_list.setCurrentRow(self.log_list.count() - 1)
    
    def clear_logs(self):
        """로그 삭제"""
        reply = QMessageBox.question(self, "로그 삭제", 
                                   "모든 로그를 삭제하시겠습니까?",
                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            logger.clear_logs()
            self.apply_filters()
            QMessageBox.information(self, "완료", "로그가 삭제되었습니다.")
    
    def export_logs(self):
        """로그 내보내기"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"barcode_logs_{timestamp}.txt"
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "로그 내보내기", default_filename, "Text Files (*.txt);;All Files (*)")
        
        if file_path:
            if logger.export_logs(file_path):
                QMessageBox.information(self, "완료", f"로그가 내보내기되었습니다:\n{file_path}")
            else:
                QMessageBox.critical(self, "오류", "로그 내보내기에 실패했습니다.")
    
    def open_log_folder(self):
        """로그 폴더 열기"""
        import os
        import subprocess
        
        log_dir = os.path.dirname(logger.log_file_path)
        
        try:
            if os.name == 'nt':  # Windows
                os.startfile(log_dir)
            elif os.name == 'posix':  # macOS, Linux
                subprocess.run(['open' if sys.platform == 'darwin' else 'xdg-open', log_dir])
        except Exception as e:
            QMessageBox.critical(self, "오류", f"로그 폴더를 열 수 없습니다:\n{e}")
    
    def closeEvent(self, event):
        """창 닫기 이벤트"""
        self.refresh_timer.stop()
        event.accept()