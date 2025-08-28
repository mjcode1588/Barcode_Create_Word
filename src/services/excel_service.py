from openpyxl import load_workbook, Workbook
from typing import List, Tuple, Optional
from src.models.product import Product
import os

class ExcelService:
    """Excel 파일 읽기/쓰기 서비스 (CRUD 지원)"""
    
    def __init__(self, file_path: str = "data/items.xlsx"):
        self.file_path = file_path
        self.categories = []  # 종류 목록 저장
        self._ensure_file_exists()
        self._load_categories()  # 시작 시 종류 목록 로드
    
    def _ensure_file_exists(self):
        """Excel 파일이 존재하지 않으면 기본 구조로 생성"""
        if not os.path.exists(self.file_path):
            self._create_default_file()
    
    def _create_default_file(self):
        """기본 Excel 파일 생성 (product와 type 시트 포함)"""
        wb = Workbook()
        
        # product 시트 생성
        product_ws = wb.active
        product_ws.title = "product"
        
        # product 시트 헤더 추가
        product_headers = ["상품명", "가격", "종류", "복사"]
        product_ws.append(product_headers)
        
        # product 시트 헤더 스타일 설정
        for col in range(1, len(product_headers) + 1):
            cell = product_ws.cell(row=1, column=col)
            cell.font = cell.font.copy(bold=True)
        
        # type 시트 생성
        type_ws = wb.create_sheet("type")
        
        # type 시트 헤더 추가
        type_headers = ["종류"]
        type_ws.append(type_headers)
        
        # type 시트 헤더 스타일 설정
        for col in range(1, len(type_headers) + 1):
            cell = type_ws.cell(row=1, column=col)
            cell.font = cell.font.copy(bold=True)
        
        # 기본 종류 데이터 추가
        default_categories = ["식품", "의류", "전자제품", "도서", "생활용품"]
        for category in default_categories:
            type_ws.append([category])
        
        # 파일 저장
        os.makedirs(os.path.dirname(self.file_path), exist_ok=True)
        wb.save(self.file_path)
        wb.close()
        print(f"기본 Excel 파일 생성: {self.file_path}")
    
    def _load_categories(self):
        """type 시트에서 종류 목록 로드"""
        try:
            wb = load_workbook(self.file_path)
            
            # type 시트가 있는지 확인
            if "type" not in wb.sheetnames:
                print("type 시트가 없습니다. 기본 종류를 사용합니다.")
                self.categories = ["식품", "의류", "전자제품", "도서", "생활용품"]
                wb.close()
                return
            
            type_ws = wb["type"]
            self.categories = []
            
            # 데이터 읽기 (2번째 행부터)
            for row in type_ws.iter_rows(min_row=2, values_only=True):
                if row[0] is not None:  # 빈 행이 아니면
                    category = str(row[0]).strip()
                    if category and category not in self.categories:
                        self.categories.append(category)
            
            wb.close()
            print(f"종류 목록 로드됨: {self.categories}")
            
        except Exception as e:
            print(f"종류 목록 로드 실패: {e}")
            self.categories = ["식품", "의류", "전자제품", "도서", "생활용품"]
    
    def get_categories(self) -> List[str]:
        """종류 목록 반환"""
        return self.categories.copy()
    
    def add_category(self, category: str) -> bool:
        """새 종류 추가"""
        try:
            if category in self.categories:
                return True  # 이미 존재하면 성공으로 처리
            
            wb = load_workbook(self.file_path)
            
            # type 시트가 없으면 생성
            if "type" not in wb.sheetnames:
                type_ws = wb.create_sheet("type")
                type_ws.append(["종류"])
            else:
                type_ws = wb["type"]
            
            # 새 종류 추가
            type_ws.append([category])
            
            wb.save(self.file_path)
            wb.close()
            
            # 메모리의 종류 목록도 업데이트
            self.categories.append(category)
            print(f"새 종류 추가됨: {category}")
            return True
            
        except Exception as e:
            print(f"종류 추가 실패: {e}")
            return False
    
    def read_products(self) -> List[Product]:
        """product 시트에서 상품 정보 읽기 (Read)"""
        try:
            wb = load_workbook(self.file_path)
            
            # product 시트 선택
            if "product" in wb.sheetnames:
                ws = wb["product"]
            else:
                # product 시트가 없으면 첫 번째 시트 사용
                ws = wb.active
                print("product 시트가 없어서 첫 번째 시트를 사용합니다.")
            
            # 헤더 읽기
            headers = []
            for cell in ws[1]:
                headers.append(cell.value)
            
            print(f"Excel 파일 컬럼: {headers}")
            
            products = []
            
            # 데이터 읽기 (2번째 행부터)
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] is None:  # 빈 행이면 건너뛰기
                    continue
                
                name = str(row[0]).strip() if row[0] else ""
                price = str(row[1]).strip() if row[1] else ""
                category = str(row[2]).strip() if row[2] else ""
                copy_flag = row[3] if len(row) > 3 else False
                
                if name and price and category:  # 필수 필드가 모두 있는 경우만
                    try:
                        product = Product(
                            name=name,
                            price=price,
                            category=category,
                            copy=copy_flag
                        )
                        products.append(product)
                    except ValueError as e:
                        print(f"상품 데이터 오류 (행 {len(products) + 2}): {e}")
                        continue
            
            wb.close()
            print(f"총 {len(products)}개 상품 로드됨")
            return products
            
        except Exception as e:
            print(f"Excel 파일 읽기 실패: {e}")
            return []
    
    def save_products(self, products: List[Product]) -> bool:
        """상품 목록을 product 시트에 저장 (Create/Update)"""
        try:
            # 기존 파일이 있으면 로드, 없으면 새로 생성
            if os.path.exists(self.file_path):
                wb = load_workbook(self.file_path)
            else:
                wb = Workbook()
            
            # product 시트 처리
            if "product" in wb.sheetnames:
                # 기존 product 시트 삭제 후 재생성
                wb.remove(wb["product"])
            
            product_ws = wb.create_sheet("product", 0)  # 첫 번째 위치에 생성
            
            # 헤더 추가
            headers = ["상품명", "가격", "종류", "복사"]
            product_ws.append(headers)
            
            # 헤더 스타일 설정
            for col in range(1, len(headers) + 1):
                cell = product_ws.cell(row=1, column=col)
                cell.font = cell.font.copy(bold=True)
            
            # 데이터 추가
            for product in products:
                product_ws.append([
                    product.name,
                    product.price,
                    product.category,
                    product.copy
                ])
            
            # type 시트가 없으면 생성
            if "type" not in wb.sheetnames:
                type_ws = wb.create_sheet("type")
                type_ws.append(["종류"])
                
                # 기본 종류 추가
                for category in self.categories:
                    type_ws.append([category])
            
            # 파일 저장
            wb.save(self.file_path)
            wb.close()
            
            print(f"Excel 파일 저장 완료: {self.file_path} ({len(products)}개 상품)")
            return True
            
        except Exception as e:
            print(f"Excel 파일 저장 실패: {e}")
            return False
    
    def add_product(self, product: Product) -> bool:
        """새 상품 추가 (Create)"""
        try:
            products = self.read_products()
            products.append(product)
            return self.save_products(products)
        except Exception as e:
            print(f"상품 추가 실패: {e}")
            return False
    
    def update_product(self, old_product: Product, new_product: Product) -> bool:
        """상품 정보 수정 (Update)"""
        try:
            products = self.read_products()
            
            # 기존 상품 찾아서 교체
            for i, existing_product in enumerate(products):
                if (existing_product.name == old_product.name and 
                    existing_product.category == old_product.category):
                    products[i] = new_product
                    break
            
            return self.save_products(products)
        except Exception as e:
            print(f"상품 수정 실패: {e}")
            return False
    
    def delete_product(self, product: Product) -> bool:
        """상품 삭제 (Delete)"""
        try:
            products = self.read_products()
            
            # 상품 찾아서 삭제
            products = [p for p in products if not (
                p.name == product.name and p.category == product.category
            )]
            
            return self.save_products(products)
        except Exception as e:
            print(f"상품 삭제 실패: {e}")
            return False
    
    def get_product_by_name_category(self, name: str, category: str) -> Optional[Product]:
        """이름과 종류로 상품 찾기"""
        try:
            products = self.read_products()
            for product in products:
                if product.name == name and product.category == category:
                    return product
            return None
        except Exception as e:
            print(f"상품 검색 실패: {e}")
            return None
    
    def get_category_counters(self, products: List[Product]) -> dict:
        """종류별 카운터 반환"""
        category_counters = {}
        for product in products:
            if product.category not in category_counters:
                category_counters[product.category] = 1
            else:
                category_counters[product.category] += 1
        return category_counters
    
    def _get_category_code(self, category: str) -> str:
        """종류명을 영문 코드로 변환"""
        category_codes = {
            "키링": "KEYRING",
            "식품": "FOOD",
            "의류": "CLOTH",
            "전자제품": "ELEC",
            "도서": "BOOK",
            "생활용품": "LIFE",
            "액세서리": "ACC",
            "장난감": "TOY",
            "화장품": "COSM",
            "스포츠": "SPORT"
        }
        
        # 등록된 코드가 있으면 사용, 없으면 영문자만 추출하거나 기본값 사용
        if category in category_codes:
            return category_codes[category]
        else:
            # 영문자만 추출
            english_only = ''.join(c for c in category if c.isalpha() and ord(c) < 128)
            if english_only:
                return english_only.upper()[:6]  # 최대 6자리
            else:
                return "ITEM"  # 기본값

    def generate_barcode_numbers(self, products: List[Product]) -> List[Tuple[str, str, str, str]]:
        """바코드 번호 생성 (한글 종류명을 영문 코드로 변환)"""
        category_counters = {}
        items = []
        
        for product in products:
            # 종류별 카운터 관리
            if product.category not in category_counters:
                category_counters[product.category] = 1
            
            # 종류의 인덱스를 사용 (1-based). 예: PPON-1 000001 -> PPON-1000001
            # 만약 종류가 아직 categories 목록에 없으면 목록에 추가해서 인덱스를 생성합니다.
            if product.category not in self.categories:
                self.categories.append(product.category)

            # 1-based 인덱스
            category_index = self.categories.index(product.category) + 1

            # 종류별 바코드 생성: PPON-{종류인덱스}{6자리순번} 형태
            item_number = str(category_counters[product.category]).zfill(6)
            barcode_number = f"PPON-{category_index}{item_number}"
            
            # 복사 플래그에 따라 개수 결정
            if product.copy:
                # True면 78개 추가
                for i in range(78):
                    items.append((product.name, product.price, product.category, barcode_number))
            else:
                # False면 1개만 추가
                items.append((product.name, product.price, product.category, barcode_number))
            
            category_counters[product.category] += 1
        
        print(f"총 {len(items)}개 라벨 생성됨")
        for category, count in category_counters.items():
            try:
                idx = self.categories.index(category) + 1
            except ValueError:
                idx = "?"
            print(f"  {category} (index:{idx}): {count-1}개")
        
        return items
    
    def backup_file(self, backup_path: str = None) -> bool:
        """Excel 파일 백업"""
        try:
            if backup_path is None:
                import datetime
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_path = f"{self.file_path}.backup_{timestamp}"
            
            import shutil
            shutil.copy2(self.file_path, backup_path)
            print(f"파일 백업 완료: {backup_path}")
            return True
            
        except Exception as e:
            print(f"파일 백업 실패: {e}")
            return False
