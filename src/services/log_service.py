#!/usr/bin/env python3
"""
로그 서비스 - 모든 애플리케이션 로그를 중앙 집중 관리
"""

import os
import sys
from datetime import datetime
from typing import List, Optional
from PyQt6.QtCore import QObject, pyqtSignal
from pathlib import Path


class LogLevel:
    """로그 레벨 상수"""
    DEBUG = "DEBUG"
    INFO = "INFO"
    WARNING = "WARNING"
    ERROR = "ERROR"
    CRITICAL = "CRITICAL"


class LogEntry:
    """로그 엔트리 클래스"""
    
    def __init__(self, level: str, message: str, module: str = "", timestamp: Optional[datetime] = None):
        self.timestamp = timestamp or datetime.now()
        self.level = level
        self.message = message
        self.module = module
    
    def __str__(self):
        return f"[{self.timestamp.strftime('%Y-%m-%d %H:%M:%S')}] [{self.level}] {self.module}: {self.message}"
    
    def to_html(self):
        """HTML 형식으로 변환"""
        color_map = {
            LogLevel.DEBUG: "#666666",
            LogLevel.INFO: "#000000", 
            LogLevel.WARNING: "#ff8c00",
            LogLevel.ERROR: "#d83b01",
            LogLevel.CRITICAL: "#a80000"
        }
        
        color = color_map.get(self.level, "#000000")
        timestamp_str = self.timestamp.strftime('%H:%M:%S')
        
        return f'<span style="color: {color};">[{timestamp_str}] [{self.level}] {self.module}: {self.message}</span>'


class LogService(QObject):
    """중앙 집중식 로그 서비스"""
    
    # 새 로그 엔트리가 추가될 때 발생하는 시그널
    log_added = pyqtSignal(LogEntry)
    
    _instance = None
    _initialized = False
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        if LogService._initialized:
            return
        
        super().__init__()
        LogService._initialized = True
        self.logs: List[LogEntry] = []
        self.max_logs = 1000  # 최대 로그 수
        self.log_file_path = self._get_log_file_path()
        
        # 시작 로그
        self.info("LogService", "로그 서비스 초기화 완료")
    
    def _get_log_file_path(self) -> str:
        """로그 파일 경로 반환"""
        if getattr(sys, 'frozen', False):
            # PyInstaller로 패키징된 경우
            base_path = os.path.dirname(sys.executable)
        else:
            # 일반 스크립트 실행
            base_path = os.path.abspath(".")
        
        log_dir = os.path.join(base_path, "logs")
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        today = datetime.now().strftime("%Y%m%d")
        return os.path.join(log_dir, f"barcode_generator_{today}.log")
    
    def _add_log(self, level: str, module: str, message: str):
        """로그 엔트리 추가"""
        entry = LogEntry(level, message, module)
        self.logs.append(entry)
        
        # 최대 로그 수 제한
        if len(self.logs) > self.max_logs:
            self.logs = self.logs[-self.max_logs:]
        
        # 파일에 저장
        self._write_to_file(entry)
        
        # 시그널 발생
        self.log_added.emit(entry)
    
    def _write_to_file(self, entry: LogEntry):
        """로그를 파일에 저장"""
        try:
            with open(self.log_file_path, 'a', encoding='utf-8') as f:
                f.write(str(entry) + '\n')
        except Exception as e:
            # 파일 저장 실패 시 콘솔에만 출력
            print(f"로그 파일 저장 실패: {e}")
    
    def debug(self, module: str, message: str):
        """디버그 로그"""
        self._add_log(LogLevel.DEBUG, module, message)
    
    def info(self, module: str, message: str):
        """정보 로그"""
        self._add_log(LogLevel.INFO, module, message)
    
    def warning(self, module: str, message: str):
        """경고 로그"""
        self._add_log(LogLevel.WARNING, module, message)
    
    def error(self, module: str, message: str):
        """에러 로그"""
        self._add_log(LogLevel.ERROR, module, message)
    
    def critical(self, module: str, message: str):
        """치명적 에러 로그"""
        self._add_log(LogLevel.CRITICAL, module, message)
    
    def get_logs(self, level_filter: Optional[str] = None, module_filter: Optional[str] = None) -> List[LogEntry]:
        """로그 목록 반환 (필터링 가능)"""
        filtered_logs = self.logs
        
        if level_filter:
            filtered_logs = [log for log in filtered_logs if log.level == level_filter]
        
        if module_filter:
            filtered_logs = [log for log in filtered_logs if module_filter.lower() in log.module.lower()]
        
        return filtered_logs
    
    def clear_logs(self):
        """모든 로그 삭제"""
        self.logs.clear()
        self.info("LogService", "로그가 삭제되었습니다")
    
    def export_logs(self, file_path: str) -> bool:
        """로그를 파일로 내보내기"""
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write("=== 바코드 라벨 생성기 로그 ===\n")
                f.write(f"내보내기 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"총 로그 수: {len(self.logs)}\n\n")
                
                for log in self.logs:
                    f.write(str(log) + '\n')
            
            self.info("LogService", f"로그 내보내기 완료: {file_path}")
            return True
        except Exception as e:
            self.error("LogService", f"로그 내보내기 실패: {e}")
            return False


# 전역 로그 서비스 인스턴스
logger = LogService()