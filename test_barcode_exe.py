#!/usr/bin/env python3
"""
exe 환경에서 바코드 생성 테스트
"""

import sys
import os
from pathlib import Path

# 현재 디렉토리를 Python 경로에 추가
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

from src.services.barcode_generator import BarcodeFileGenerator

def test_barcode_generation():
    """바코드 생성 테스트"""
    print("바코드 생성 테스트 시작...")
    
    # 테스트 데이터
    test_products = [
        ("테스트상품1", "1000", "식품", "1000001"),
        ("테스트상품2", "2000", "음료", "2000002"),
        ("테스트상품3", "3000", "과자", "3000003"),
    ]
    
    try:
        # 바코드 파일 생성기 초기화
        barcode_generator = BarcodeFileGenerator("test_barcodes")
        print(f"바코드 출력 디렉토리: {barcode_generator.output_dir}")
        
        # 바코드 생성
        generated_files = barcode_generator.generate_barcodes_for_products(test_products)
        
        print(f"생성된 바코드 파일 수: {len(generated_files)}")
        
        for code, filepath in generated_files.items():
            if os.path.exists(filepath):
                file_size = os.path.getsize(filepath)
                print(f"✓ {code}: {filepath} ({file_size} bytes)")
            else:
                print(f"✗ {code}: 파일이 생성되지 않음 - {filepath}")
        
        # 정리
        barcode_generator.cleanup()
        print("테스트 완료!")
        
    except Exception as e:
        print(f"테스트 실패: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_barcode_generation()