import os
from barcode import Code128
from barcode.writer import ImageWriter
from typing import List, Tuple, Optional, Dict
from io import BytesIO
from PIL import Image


class BarcodeGenerator:
    """GS1-128 바코드 생성 클래스 (메모리 기반) - MM 단위 통일"""

    def __init__(self, options: Optional[Dict] = None):
        if options is None:
            options = {}

        # MM 단위로 입력받은 값들을 바코드 라이브러리에 맞게 변환
        self.writer_options = self._convert_mm_to_barcode_units(options)

    def _convert_mm_to_barcode_units(self, options: Dict) -> Dict:
        """MM 단위 입력값을 바코드 라이브러리 단위로 변환 (안전한 범위 내에서 적용)"""
        # 사용자 설정값 (MM 단위)
        module_width_mm = options.get("module_width", 0.3)
        module_height_mm = options.get("module_height", 15.0)
        quiet_zone_mm = options.get("quiet_zone", 3.0)
        text_distance_mm = options.get("text_distance", 5.0)

        # 바코드 라이브러리 안전 범위 내에서만 적용
        # module_width와 quiet_zone은 라이브러리 기본값 사용 (단위 문제 회피)
        # module_height와 text_distance는 MM 단위로 사용자 값 적용

        return {
            "module_width": 0.35,  # 더 넓은 바코드를 위해 증가 (0.2 → 0.35)
            "module_height": max(
                10.0, min(25.0, module_height_mm)
            ),  # 사용자 값 적용 (10-25mm 범위)
            "quiet_zone": 6.5,  # 라이브러리 기본값 (안정성 우선)
            "text_distance": max(
                3.0, min(8.0, text_distance_mm)
            ),  # 사용자 값 적용 (3-8mm 범위)
            "font_size": options.get("font_size", 10),
            "dpi": options.get("dpi", 300),
        }

    def generate_barcode_in_memory(self, code: str, category: str) -> BytesIO:
        """바코드 이미지를 메모리에 생성"""
        try:
            # 바코드 데이터가 ASCII 문자만 포함하는지 확인
            try:
                code.encode("ascii")
            except UnicodeEncodeError:
                print(f"바코드 데이터에 ASCII가 아닌 문자가 포함됨: {code}")
                # ASCII가 아닌 문자를 제거하거나 변환
                code = "".join(c for c in code if ord(c) < 128)
                print(f"ASCII 문자만 추출: {code}")

            if not code:
                print(f"유효한 바코드 데이터가 없음")
                return None

            # GS1-128 바코드 생성 (안정적인 기본 설정 사용)
            writer = ImageWriter(format="PNG")

            # Code128로 GS1-128 형식 생성
            barcode = Code128(code, writer=writer)

            # 메모리에 이미지 생성 (안정적인 기본 옵션 사용)
            img_buffer = BytesIO()
            barcode.write(img_buffer, self.writer_options)
            img_buffer.seek(0)

            print(f"바코드 생성 완료: {code} ({category})")
            return img_buffer

        except Exception as e:
            print(f"바코드 생성 실패: {code} - {e}")
            return None

    def generate_barcodes_for_products(
        self, products: List[Tuple[str, str, str, str]]
    ) -> dict:
        """상품 목록에 대한 바코드 생성 (메모리에 저장)"""
        barcode_images = {}

        for name, price, category, code in products:
            if code not in barcode_images:
                img_buffer = self.generate_barcode_in_memory(code, category)
                if img_buffer:
                    barcode_images[code] = img_buffer

        return barcode_images
