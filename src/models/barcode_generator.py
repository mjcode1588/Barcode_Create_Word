import os
from barcode import Code128
from barcode.writer import ImageWriter
from typing import List, Tuple

class BarcodeGenerator:
    """GS1-128 바코드 생성 클래스"""
    
    def __init__(self, output_dir: str = "barcodes"):
        self.output_dir = output_dir
        self._ensure_output_dir()
        
        # 바코드 설정
        self.barcode_width = 1.2
        self.barcode_height = 0.6
        self.dpi = 300
        self.module_width = 0.2
        self.module_height = 15.0
        self.quiet_zone = 6.5
        self.text_distance = 5.0
        self.font_size = 10
    
    def _ensure_output_dir(self):
        """출력 디렉토리 생성"""
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
    
    def generate_barcode(self, code: str, category: str) -> str:
        """바코드 이미지 생성"""
        filename = os.path.join(self.output_dir, f"{code}.png")
        
        if os.path.exists(filename):
            return filename
        
        try:
            # GS1-128 바코드 생성 (고화질 설정)
            writer = ImageWriter()
            writer.dpi = self.dpi
            writer.module_width = self.module_width
            writer.module_height = self.module_height
            writer.quiet_zone = self.quiet_zone
            writer.text_distance = self.text_distance
            writer.font_size = self.font_size
            
            # Code128로 GS1-128 형식 생성
            barcode = Code128(code, writer=writer)
            barcode.save(os.path.join(self.output_dir, code))
            
            print(f"바코드 생성 완료: {code} ({category})")
            return filename
            
        except Exception as e:
            print(f"바코드 생성 실패: {code} - {e}")
            return ""
    
    def generate_barcodes_for_products(self, products: List[Tuple[str, str, str, str]]) -> List[str]:
        """상품 목록에 대한 바코드 생성"""
        unique_codes = set()
        generated_files = []
        
        for name, price, category, code in products:
            if code not in unique_codes:
                unique_codes.add(code)
                filename = self.generate_barcode(code, category)
                if filename:
                    generated_files.append(filename)
        
        return generated_files
    
    def cleanup(self):
        """바코드 파일 정리"""
        if os.path.exists(self.output_dir):
            import shutil
            shutil.rmtree(self.output_dir)
            print(f"바코드 디렉토리 정리 완료: {self.output_dir}")
