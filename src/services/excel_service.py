from openpyxl import load_workbook, Workbook
from typing import List, Tuple, Optional
from src.models.product import Product
import os

class ExcelService:
    """Excel 파일 읽기/쓰기 서비스 (CRUD 지원)"""
    
    def __init__(self, file_path: str = "data/items.xlsx"):
        self.file_path = file_path
        self.categories = []  # TYPE 목록 저장
        self.categories_dict = {}
        self._ensure_file_exists()
        self._load_categories()  # 시작 시 TYPE 목록 로드
    
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
        product_headers = ["PRODUCT", "PRICE", "TYPE", "PRODUCT_ID"]
        product_ws.append(product_headers)
        
        # product 시트 헤더 스타일 설정
        for col in range(1, len(product_headers) + 1):
            cell = product_ws.cell(row=1, column=col)
            cell.font = cell.font.copy(bold=True)
        
        # type 시트 생성
        type_ws = wb.create_sheet("type")
        
        # type 시트 헤더 추가
        type_headers = ["TYPE", "TYPE_ID"]
        type_ws.append(type_headers)
        
        # type 시트 헤더 스타일 설정
        for col in range(1, len(type_headers) + 1):
            cell = type_ws.cell(row=1, column=col)
            cell.font = cell.font.copy(bold=True)


        # 기본 TYPE 데이터 추가
        default_categories = ["폰스트랩", "리본 키링", "미니 키링", "키링", "팔찌", "꽃갈피", "모양", "부착"]
        for idx, type_name in enumerate(default_categories):
            type_ws.append([type_name, idx])
        
        # 파일 저장
        os.makedirs(os.path.dirname(self.file_path), exist_ok=True)
        wb.save(self.file_path)
        wb.close()
        print(f"기본 Excel 파일 생성: {self.file_path}")
    
    def _load_categories(self):
        """type 시트에서 TYPE 목록 로드"""
        try:
            wb = load_workbook(self.file_path)
            
            # type 시트가 있는지 확인
            if "type" not in wb.sheetnames:
                print("type 시트가 없습니다. 기본 TYPE를 사용합니다.")
                self.categories = ["폰스트랩", "리본 키링", "미니 키링", "키링", "팔찌", "꽃갈피", "모양", "부착"]
                self.categories_dict = dict(self.categories, zip(range(len(self.categories))))
                wb.close()
                return
            
            type_ws = wb["type"]
            
            # 기존 메모리의 카테고리와 파일의 카테고리를 병합
            # 이렇게 하면 UI에서 임시로 추가된 카테고리가 사라지지 않음
            file_categories = []
            categories_idx = []
            for row in type_ws.iter_rows(min_row=2, values_only=True):
                if row[0] is not None:
                    type_name = str(row[0]).strip()
                    if type_name:
                        file_categories.append(type_name)
                if row[1] is not None:
                    idx = row[1]
                    if type_name:
                        categories_idx.append(idx)
            #  파일 카테고리로 덮어 쓰기
             
            self.categories = file_categories
            self.categories_dict = dict(zip(categories_idx,file_categories))

            wb.close()
            print(f"TYPE 목록 로드됨: {self.categories}")
            print(f"TYPE DICT 로드됨: {self.categories_idx}")
            
        except Exception as e:
            print(f"TYPE 목록 로드 실패: {e}")
            self.categories = ["폰스트랩", "리본 키링", "미니 키링", "키링", "팔찌", "꽃갈피", "모양", "부착"]
    
    def get_categories(self) -> List[str]:
        """TYPE 목록 반환"""
        return self.categories.copy()
    
    def is_type_name_in_use(self, type_name: str) -> bool:
        """해당 TYPE가 상품에서 사용 중인지 확인"""
        products = self.read_products()
        return any(p.type_name == type_name for p in products)

    def update_type_name(self, old_type_name: str, new_type_name: str) -> bool:
        """TYPE 수정 (연관된 모든 상품 정보 포함)"""
        if not new_type_name or new_type_name in self.categories:
            return False
        
        try:
            wb = load_workbook(self.file_path)
            
            # 1. type 시트에서 TYPE 수정
            type_ws = wb["type"]
            for row in type_ws.iter_rows(min_row=2):
                if row[0].value == old_type_name:
                    row[0].value = new_type_name
                    break
            
            # 2. product 시트에서 해당 TYPE를 사용하는 모든 상품 수정
            if "product" in wb.sheetnames:
                product_ws = wb["product"]
                for row in product_ws.iter_rows(min_row=2):
                    if row[2].value == old_type_name:
                        row[2].value = new_type_name
            
            wb.save(self.file_path)
            wb.close()
            
            # 3. 메모리의 TYPE 목록 업데이트
            self.categories = [new_type_name if c == old_type_name else c for c in self.categories]
            print(f"TYPE 수정 완료: '{old_type_name}' -> '{new_type_name}'")
            return True
            
        except Exception as e:
            print(f"TYPE 수정 실패: {e}")
            return False

    def delete_type_name(self, type_name: str) -> bool:
        """TYPE 삭제"""
        if self.is_type_name_in_use(type_name):
            print(f"'{type_name}' TYPE는 현재 사용 중이므로 삭제할 수 없습니다.")
            return False
        
        try:
            wb = load_workbook(self.file_path)
            type_ws = wb["type"]
            
            # 삭제할 행 찾기
            row_to_delete = None
            for row_idx, row in enumerate(type_ws.iter_rows(min_row=2), start=2):
                if row[0].value == type_name:
                    row_to_delete = row_idx
                    break
            
            # 행 삭제
            if row_to_delete:
                type_ws.delete_rows(row_to_delete)
            
            wb.save(self.file_path)
            wb.close()
            
            # 메모리에서 TYPE 삭제
            if type_name in self.categories:
                self.categories.remove(type_name)
            
            print(f"TYPE 삭제 완료: '{type_name}'")
            return True
            
        except Exception as e:
            print(f"TYPE 삭제 실패: {e}")
            return False

    def add_type_name(self, type_name: str) -> bool:
        """새 TYPE 추가"""
        try:
            if type_name in self.categories:
                return True  # 이미 존재하면 성공으로 처리
            
            wb = load_workbook(self.file_path)
            
            # type 시트가 없으면 생성
            if "type" not in wb.sheetnames:
                type_ws = wb.create_sheet("type")
                type_ws.append(["TYPE"])
            else:
                type_ws = wb["type"]
            
            # 새 TYPE 추가
            type_ws.append([type_name])
            
            wb.save(self.file_path)
            wb.close()
            
            # 메모리의 TYPE 목록도 업데이트
            self.categories.append(type_name)
            print(f"새 TYPE 추가됨: {type_name}")
            return True
            
        except Exception as e:
            print(f"TYPE 추가 실패: {e}")
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
                type_id = str(row[2]).strip() if row[1] else ""
                type_name = str(self.categories_dict[row[2]]).strip() if row[2] else ""
                product_id = row[3] if len(row) > 3 else ""
                
                if name and price and type_name:  # 필수 필드가 모두 있는 경우만
                    try:
                        product = Product(
                            name=name,
                            price=price,
                            type_name=type_name,
                            type_id= type_id,
                            product_id= product_id
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
    
    def save_products(self, products: List[Product], temp_file_path:str) -> bool:
        """상품 목록을 product 시트에 저장 (Create/Update)"""
        try:
            # 기존 파일이 있으면 로드, 없으면 새로 생성
            if os.path.exists(temp_file_path):
                wb = load_workbook(temp_file_path)
            else:
                wb = Workbook()
            
            # product 시트 처리
            if "product" in wb.sheetnames:
                # 기존 product 시트 삭제 후 재생성
                wb.remove(wb["product"])
            
            product_ws = wb.create_sheet("product", 0)  # 첫 번째 위치에 생성
            
            # 헤더 추가
            headers = ["상품명", "가격", "TYPE", "복사"]
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
                    product.type_name,
                    product.copy
                ])
            
            # type 시트 처리: 기존 시트 삭제 후 현재 TYPE 목록으로 새로 생성
            if "type" in wb.sheetnames:
                wb.remove(wb["type"])
            
            type_ws = wb.create_sheet("type")
            
            # 헤더 추가
            type_headers = ["TYPE"]
            type_ws.append(type_headers)
            
            # 헤더 스타일 설정
            cell = type_ws.cell(row=1, column=1)
            cell.font = cell.font.copy(bold=True)
            
            # 현재 TYPE 목록 저장
            for type_name in self.categories:
                type_ws.append([type_name])
            

            if "Sheet" in wb.sheetnames:
                wb.remove(wb["Sheet"])

            # 파일 저장
            wb.save(temp_file_path)
            wb.close()
            
            print(f"Excel 파일 저장 완료: {temp_file_path} ({len(products)}개 상품)")
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
                    existing_product.type_name == old_product.type_name):
                    products[i] = new_product
                    break
            
            return self.save_products(products, self.file_path)
        except Exception as e:
            print(f"상품 수정 실패: {e}")
            return False
    
    def delete_product(self, product: Product) -> bool:
        """상품 삭제 (Delete)"""
        try:
            products = self.read_products()
            
            # 상품 찾아서 삭제
            products = [p for p in products if not (
                p.name == product.name and p.type_name == product.type_name
            )]
            
            return self.save_products(products)
        except Exception as e:
            print(f"상품 삭제 실패: {e}")
            return False
    
    def get_product_by_name_type_name(self, name: str, type_name: str) -> Optional[Product]:
        """이름과 TYPE로 상품 찾기"""
        try:
            products = self.read_products()
            for product in products:
                if product.name == name and product.type_name == type_name:
                    return product
            return None
        except Exception as e:
            print(f"상품 검색 실패: {e}")
            return None
    
    def get_type_name_counters(self, products: List[Product]) -> dict:
        """TYPE별 카운터 반환"""
        type_name_counters = {}
        for product in products:
            if product.type_name not in type_name_counters:
                type_name_counters[product.type_name] = 1
            else:
                type_name_counters[product.type_name] += 1
        return type_name_counters
    
    def _get_type_name_code(self, type_name: str) -> str:
        """TYPE명을 영문 코드로 변환"""
        type_name_codes = {
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
        if type_name in type_name_codes:
            return type_name_codes[type_name]
        else:
            # 영문자만 추출
            english_only = ''.join(c for c in type_name if c.isalpha() and ord(c) < 128)
            if english_only:
                return english_only.upper()[:6]  # 최대 6자리
            else:
                return "ITEM"  # 기본값

    def generate_barcode_numbers(self, products: List[Product]) -> List[Tuple[str, str, str, str]]:
        """바코드 번호 생성 (한글 종류명을 영문 코드로 변환) 및 새 종류 자동 저장"""
        # 상품 목록에서 새로운 종류를 찾아 파일에 미리 저장
        product_categories = sorted(list(set(p.type_name for p in products)))
        new_categories = [cat for cat in product_categories if cat not in self.categories]

        if new_categories:
            try:
                wb = load_workbook(self.file_path)
                
                if "type" not in wb.sheetnames:
                    type_ws = wb.create_sheet("type")
                    type_ws.append(["TYPE"])
                    # 헤더 스타일 설정
                    cell = type_ws.cell(row=1, column=1)
                    cell.font = cell.font.copy(bold=True)
                else:
                    type_ws = wb["type"]
                
                for type_name in new_categories:
                    type_ws.append([type_name])
                
                wb.save(self.file_path)
                wb.close()
                
                # 메모리 내 종류 목록 업데이트
                self.categories.extend(new_categories)
                print(f"새 종류 자동 저장됨: {new_categories}")

            except Exception as e:
                print(f"새 종류 자동 저장 실패: {e}")
                # 실패하더라도 메모리에 추가하여 바코드 생성은 계속 진행
                self.categories.extend(new_categories)

        type_name_counters = {}
        items = []
        
        for product in products:           
            # 종류별 바코드 생성: PPON-{종류인덱스}{6자리순번} 형태
            item_number = str(product.product_id).zfill(6)
            barcode_number = f"PPON-{product.type_id}{item_number}"
            

            # 고쳐야되는데, word의 내용을 받아와야됨스
            for _ in range(78):
                items.append((product.name, product.price, product.type_name, barcode_number))
            else:
                # False면 1개만 추가
                items.append((product.name, product.price, product.type_name, barcode_number))
            
            type_name_counters[product.type_name] += 1
        
        print(f"총 {len(items)}개 라벨 생성됨")
        for type_name, count in type_name_counters.items():
            try:
                idx = self.categories.index(type_name) + 1
            except ValueError:
                idx = "?"
            print(f"  {type_name} (index:{idx}): {count-1}개")
        
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
