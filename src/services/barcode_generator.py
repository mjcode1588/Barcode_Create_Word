import os
from barcode import Code128
from barcode.writer import ImageWriter
from typing import List, Tuple
from io import BytesIO
from PIL import Image

class BarcodeGenerator:
    """GS1-128 바코드 생성 클래스 (메모리 기반)"""
    
    def __init__(self):
        # 바코드 설정
        self.barcode_width = 1.2
        self.barcode_height = 0.6
        self.dpi = 300
        self.module_width = 0.2
        self.module_height = 15.0
        self.quiet_zone = 6.5
        self.text_distance = 5.0
        self.font_size = 10
    
    def generate_barcode_in_memory(self, code: str, category: str) -> BytesIO:
        """바코드 이미지를 메모리에 생성"""
        try:
            # 바코드 데이터가 ASCII 문자만 포함하는지 확인
            try:
                code.encode('ascii')
            except UnicodeEncodeError:
                print(f"바코드 데이터에 ASCII가 아닌 문자가 포함됨: {code}")
                # ASCII가 아닌 문자를 제거하거나 변환
                code = ''.join(c for c in code if ord(c) < 128)
                print(f"ASCII 문자만 추출: {code}")
            
            if not code:
                print(f"유효한 바코드 데이터가 없음")
                return None
            
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
            
            # 메모리에 이미지 생성
            img_buffer = BytesIO()
            barcode.write(img_buffer)
            img_buffer.seek(0)
            
            print(f"바코드 생성 완료: {code} ({category})")
            return img_buffer
            
        except Exception as e:
            print(f"바코드 생성 실패: {code} - {e}")
            return None
    
    def generate_barcodes_for_products(self, products: List[Tuple[str, str, str, str]]) -> dict:
        """상품 목록에 대한 바코드 생성 (메모리에 저장)"""
        barcode_images = {}
        
        for name, price, category, code in products:
            if code not in barcode_images:
                img_buffer = self.generate_barcode_in_memory(code, category)
                if img_buffer:
                    barcode_images[code] = img_buffer
        
        return barcode_images
