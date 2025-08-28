import os
import shutil
from typing import List

class FileService:
    """파일 관리 서비스"""
    
    def __init__(self, base_dir: str = "."):
        self.base_dir = base_dir
        self.template_dir = os.path.join(base_dir, "templates")
        self.data_dir = os.path.join(base_dir, "data")
        self.output_dir = os.path.join(base_dir, "output")
    
    def ensure_directories(self):
        """필요한 디렉토리들 생성"""
        directories = [self.template_dir, self.data_dir, self.output_dir]
        for directory in directories:
            if not os.path.exists(directory):
                os.makedirs(directory)
                print(f"디렉토리 생성: {directory}")
    
    def get_template_path(self, template_name: str = "3677.docx") -> str:
        """템플릿 파일 경로 반환"""
        # 먼저 templates 디렉토리에서 찾기
        template_path = os.path.join(self.template_dir, template_name)
        if os.path.exists(template_path):
            return template_path
        
        # 현재 디렉토리에서 찾기
        current_path = os.path.join(self.base_dir, template_name)
        if os.path.exists(current_path):
            return current_path
        
        # 상위 디렉토리에서 찾기
        parent_path = os.path.join(os.path.dirname(self.base_dir), template_name)
        if os.path.exists(parent_path):
            return parent_path
        
        raise FileNotFoundError(f"템플릿 파일을 찾을 수 없습니다: {template_name}")
    
    def get_data_path(self, data_name: str = "items.xlsx") -> str:
        """데이터 파일 경로 반환"""
        # 먼저 data 디렉토리에서 찾기
        data_path = os.path.join(self.data_dir, data_name)
        if os.path.exists(data_path):
            return data_path
        
        # 현재 디렉토리에서 찾기
        current_path = os.path.join(self.base_dir, data_name)
        if os.path.exists(current_path):
            return current_path
        
        # 상위 디렉토리에서 찾기
        parent_path = os.path.join(os.path.dirname(self.base_dir), data_name)
        if os.path.exists(parent_path):
            return parent_path
        
        raise FileNotFoundError(f"데이터 파일을 찾을 수 없습니다: {data_name}")
    
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
