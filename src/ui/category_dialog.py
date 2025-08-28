from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QListWidget, 
                               QPushButton, QMessageBox, QInputDialog, QListWidgetItem)
from PyQt6.QtCore import pyqtSignal
from typing import List

from src.services.excel_service import ExcelService

class CategoryDialog(QDialog):
    """종류 관리 다이얼로그"""
    
    # 종류 목록이 변경되었음을 알리는 신호
    categories_updated = pyqtSignal()
    
    def __init__(self, excel_service: ExcelService, parent=None):
        super().__init__(parent)
        self.excel_service = excel_service
        
        self.setWindowTitle("종류 관리")
        self.setMinimumWidth(350)
        
        self.setup_ui()
        self.load_categories()
    
    def setup_ui(self):
        """UI 구성"""
        layout = QVBoxLayout(self)
        
        self.category_list = QListWidget()
        self.category_list.itemDoubleClicked.connect(self.edit_category)
        layout.addWidget(self.category_list)
        
        button_layout = QHBoxLayout()
        
        self.add_button = QPushButton("추가")
        self.add_button.clicked.connect(self.add_category)
        
        self.edit_button = QPushButton("수정")
        self.edit_button.clicked.connect(self.edit_category)
        
        self.delete_button = QPushButton("삭제")
        self.delete_button.clicked.connect(self.delete_category)
        
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.edit_button)
        button_layout.addWidget(self.delete_button)
        
        layout.addLayout(button_layout)
    
    def load_categories(self):
        """Excel 서비스에서 종류 목록을 불러와 리스트에 표시"""
        self.category_list.clear()
        categories = self.excel_service.get_categories()
        self.category_list.addItems(categories)
    
    def add_category(self):
        """새 종류 추가"""
        text, ok = QInputDialog.getText(self, "새 종류 추가", "추가할 종류 이름을 입력하세요:")
        
        if ok and text:
            text = text.strip()
            if text in self.excel_service.get_categories():
                QMessageBox.warning(self, "중복 오류", "이미 존재하는 종류입니다.")
                return
            
            if self.excel_service.add_category(text):
                self.load_categories()
                self.categories_updated.emit()
                QMessageBox.information(self, "성공", f"'{text}' 종류가 추가되었습니다.")
            else:
                QMessageBox.critical(self, "오류", "종류 추가에 실패했습니다.")
    
    def edit_category(self):
        """선택된 종류 수정"""
        selected_item = self.category_list.currentItem()
        if not selected_item:
            QMessageBox.warning(self, "선택 오류", "수정할 종류를 선택하세요.")
            return
        
        old_category = selected_item.text()
        
        new_category, ok = QInputDialog.getText(self, "종류 수정", 
                                                f"'{old_category}'의 새 이름을 입력하세요:",
                                                text=old_category)
        
        if ok and new_category:
            new_category = new_category.strip()
            if new_category == old_category:
                return
            
            if new_category in self.excel_service.get_categories():
                QMessageBox.warning(self, "중복 오류", "이미 존재하는 종류 이름입니다.")
                return
            
            if self.excel_service.update_category(old_category, new_category):
                self.load_categories()
                self.categories_updated.emit()
                QMessageBox.information(self, "성공", f"'{old_category}'가 '{new_category}'(으)로 수정되었습니다.")
            else:
                QMessageBox.critical(self, "오류", "종류 수정에 실패했습니다.")
    
    def delete_category(self):
        """선택된 종류 삭제"""
        selected_item = self.category_list.currentItem()
        if not selected_item:
            QMessageBox.warning(self, "선택 오류", "삭제할 종류를 선택하세요.")
            return
        
        category_to_delete = selected_item.text()
        
        # 해당 종류가 사용 중인지 확인
        if self.excel_service.is_category_in_use(category_to_delete):
            QMessageBox.warning(self, "삭제 불가", 
                                f"'{category_to_delete}' 종류는 현재 하나 이상의 상품에서 사용 중이므로 삭제할 수 없습니다.")
            return
        
        reply = QMessageBox.question(self, "삭제 확인",
                                     f"'{category_to_delete}' 종류를 정말 삭제하시겠습니까?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            if self.excel_service.delete_category(category_to_delete):
                self.load_categories()
                self.categories_updated.emit()
                QMessageBox.information(self, "성공", f"'{category_to_delete}' 종류가 삭제되었습니다.")
            else:
                QMessageBox.critical(self, "오류", "종류 삭제에 실패했습니다.")
