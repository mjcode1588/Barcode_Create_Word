import os
import shutil
from typing import List
import sys

def get_base_path():
    """
    실행 환경(개발용 또는 PyInstaller 번들)에 따라 기본 경로를 반환합니다.
    """
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # PyInstaller로 번들된 경우, _MEIPASS는 임시 폴더의 경로입니다.
        return sys._MEIPASS
    else:
        # 일반적인 개발 환경인 경우, 이 파일(src/services/file_service.py)을 기준으로
        # 프로젝트 루트는 두 단계 위입니다.
        return os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

class FileService:
    """파일 관리 서비스"""
    
    def __init__(self):
        self.base_dir = get_base_path()
        self.template_dir = os.path.join(self.base_dir, "templates")
        self.data_dir = os.path.join(self.base_dir, "data")
        # 출력 디렉토리는 번들 내부가 아닌 현재 작업 디렉토리에 생성합니다.
        self.output_dir = os.path.join(os.path.abspath("."), "output")
    
    def ensure_directories(self):
        """출력 디렉토리 생성 확인"""
        # data 및 templates 디렉토리는 번들에 포함되므로 출력 디렉토리만 확인합니다.
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
            print(f"디렉토리 생성: {self.output_dir}")
    
    def get_template_path(self, template_name: str = "3677.docx") -> str:
        """템플릿 파일 경로 반환"""
        template_path = os.path.join(self.template_dir, template_name)
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"템플릿 파일을 찾을 수 없습니다: {template_path}")
        return template_path
    
    def get_data_path(self, data_name: str = "items.xlsx") -> str:
        """데이터 파일 경로 반환"""
        data_path = os.path.join(self.data_dir, data_name)
        if not os.path.exists(data_path):
            # 데이터 파일이 없을 경우, 기본 파일을 생성하도록 유도할 수 있습니다.
            # 여기서는 ExcelService가 이 역할을 하므로 경로만 반환합니다.
            # raise FileNotFoundError(f"데이터 파일을 찾을 수 없습니다: {data_path}")
            pass
        return data_path
    
    def cleanup_output(self):
        """출력 디렉토리 정리"""
        if os.path.exists(self.output_dir):
            shutil.rmtree(self.output_dir)
            print(f"출력 디렉토리 정리 완료: {self.output_dir}")
    
    def list_output_files(self) -> List[str]:
        """출력 파일 목록 반환"""
        if not os.path.exists(self.output_dir):
            return []
        
        files = []
        for file in os.listdir(self.output_dir):
            if file.endswith('.docx'):
                files.append(file)
        
        return sorted(files)
    
    def open_output_directory(self):
        """출력 디렉토리 열기"""
        if os.path.exists(self.output_dir):
            os.startfile(self.output_dir)  # Windows
        else:
            print("출력 디렉토리가 존재하지 않습니다.")
    
    def get_file_size(self, file_path: str) -> str:
        """파일 크기를 읽기 쉬운 형태로 반환"""
        if not os.path.exists(file_path):
            return "0 B"
        
        size = os.path.getsize(file_path)
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024.0:
                return f"{size:.1f} {unit}"
            size /= 1024.0
        return f"{size:.1f} TB"