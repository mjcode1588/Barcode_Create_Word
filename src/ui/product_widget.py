from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
                               QLabel, QLineEdit, QCheckBox, QPushButton, QGroupBox, QComboBox)
from PyQt6.QtCore import pyqtSignal, Qt
from PyQt6.QtGui import QValidator, QIntValidator
from src.models.product import Product
from typing import List

class PriceValidator(QValidator):
    """가격 입력 검증기 (숫자만 허용)"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
    
    def validate(self, input_str, pos):
        if not input_str:
            return QValidator.State.Acceptable, input_str, pos
        
        # 숫자와 쉼표만 허용
        if all(c.isdigit() or c == ',' for c in input_str):
            return QValidator.State.Acceptable, input_str, pos
        else:
            return QValidator.State.Invalid, input_str, pos

class ProductWidget(QWidget):
    """상품 정보 입력 위젯"""
    
    productAdded = pyqtSignal(Product)
    productUpdated = pyqtSignal(Product)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_product = None
        self.categories = []  # 종류 목록
        self.setup_ui()
    
    def setup_ui(self):
        """UI 구성"""
        layout = QVBoxLayout()
        
        # 제목
        title_label = QLabel("상품 정보 입력")
        title_label.setProperty("class", "title")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)
        
        # 상품 정보 입력 그룹
        product_group = QGroupBox("상품 정보")
        product_layout = QGridLayout()
        
        # 상품명
        self.name_label = QLabel("상품명:")
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("상품명을 입력하세요")
        product_layout.addWidget(self.name_label, 0, 0)
        product_layout.addWidget(self.name_edit, 0, 1)
        
        # 가격
        self.price_label = QLabel("가격:")
        self.price_edit = QLineEdit()
        self.price_edit.setPlaceholderText("숫자만 입력 가능")
        self.price_edit.setValidator(PriceValidator())
        product_layout.addWidget(self.price_label, 1, 0)
        product_layout.addWidget(self.price_edit, 1, 1)
        
        # 종류 (콤보박스로 변경)
        self.category_label = QLabel("종류:")
        self.category_combo = QComboBox()
        self.category_combo.setEditable(True)  # 직접 입력도 가능
        self.category_combo.setPlaceholderText("종류를 선택하거나 입력하세요")
        product_layout.addWidget(self.category_label, 2, 0)
        product_layout.addWidget(self.category_combo, 2, 1)
        
        # 복사 여부
        self.copy_checkbox = QCheckBox("복사 (78개 생성)")
        self.copy_checkbox.setToolTip("체크하면 한 페이지에 78개의 라벨을 생성합니다")
        product_layout.addWidget(self.copy_checkbox, 3, 0, 1, 2)
        
        product_group.setLayout(product_layout)
        layout.addWidget(product_group)
        
        # 버튼 그룹
        button_layout = QHBoxLayout()
        
        self.add_button = QPushButton("상품 추가")
        self.add_button.clicked.connect(self.add_product)
        
        self.update_button = QPushButton("상품 수정")
        self.update_button.clicked.connect(self.update_product)
        self.update_button.setEnabled(False)
        
        self.clear_button = QPushButton("입력 초기화")
        self.clear_button.clicked.connect(self.clear_inputs)
        
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.update_button)
        button_layout.addWidget(self.clear_button)
        
        layout.addLayout(button_layout)
        
        # 상태 표시
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("color: #666666; font-style: italic;")
        layout.addWidget(self.status_label)
        
        self.setLayout(layout)
    
    def set_categories(self, categories: List[str]):
        """종류 목록 설정"""
        self.categories = categories
        self.category_combo.clear()
        self.category_combo.addItems(categories)
        print(f"종류 목록 설정됨: {categories}")
    
    def add_product(self):
        """상품 추가"""
        try:
            product = self._create_product_from_inputs()
            if product:
                self.productAdded.emit(product)
                self.clear_inputs()
                self.show_status("상품이 추가되었습니다.", "success")
        except ValueError as e:
            self.show_status(f"오류: {str(e)}", "error")
    
    def update_product(self):
        """상품 수정"""
        if not self.current_product:
            return
        
        try:
            updated_product = self._create_product_from_inputs()
            if updated_product:
                self.productUpdated.emit(updated_product)
                self.clear_inputs()
                self.current_product = None
                self.add_button.setEnabled(True)
                self.update_button.setEnabled(False)
                self.show_status("상품이 수정되었습니다.", "success")
        except ValueError as e:
            self.show_status(f"오류: {str(e)}", "error")
    
    def clear_form(self):
        """입력 폼 초기화 (내부 로직 및 외부 호출용)"""
        self.name_edit.clear()
        self.price_edit.clear()
        self.category_combo.setCurrentText("")
        self.copy_checkbox.setChecked(False)
        self.current_product = None
        self.add_button.setEnabled(True)
        self.update_button.setEnabled(False)
        self.status_label.setText("")

    def clear_inputs(self):
        """입력 필드 초기화 (사용자 액션)"""
        self.clear_form()
        self.show_status("입력 필드가 초기화되었습니다.", "info")
    
    def _create_product_from_inputs(self) -> Product:
        """입력 필드에서 Product 객체 생성"""
        name = self.name_edit.text().strip()
        price = self.price_edit.text().strip()
        category = self.category_combo.currentText().strip()
        copy = self.copy_checkbox.isChecked()
        
        if not name:
            raise ValueError("상품명을 입력해주세요.")
        
        if not price:
            raise ValueError("가격을 입력해주세요.")
        
        if not category:
            raise ValueError("종류를 입력해주세요.")
        
        return Product(name=name, price=price, category=category, copy=copy)
    
    def edit_product(self, product: Product):
        """상품 정보로 입력 필드 채우기 (수정 모드)"""
        self.current_product = product
        self.name_edit.setText(product.name)
        self.price_edit.setText(product.price)
        self.category_combo.setCurrentText(product.category)
        self.copy_checkbox.setChecked(product.copy)
        
        self.add_button.setEnabled(False)
        self.update_button.setEnabled(True)
        self.show_status("상품 정보를 수정하세요.", "info")
    
    def show_status(self, message: str, status_type: str = "info"):
        """상태 메시지 표시"""
        self.status_label.setText(message)
        
        # 상태에 따른 스타일 적용
        if status_type == "success":
            self.status_label.setStyleSheet("color: #107c10; font-weight: bold;")
        elif status_type == "error":
            self.status_label.setStyleSheet("color: #d83b01; font-weight: bold;")
        else:
            self.status_label.setStyleSheet("color: #666666; font-style: italic;")
    
    def is_input_valid(self) -> bool:
        """입력 유효성 검사"""
        try:
            self._create_product_from_inputs()
            return True
        except ValueError:
            return False