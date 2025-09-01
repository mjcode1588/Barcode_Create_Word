import os
import sys
from PyQt6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                               QSplitter, QTableWidget, QTableWidgetItem, QPushButton,
                               QLabel, QProgressBar, QMessageBox, QFileDialog, QMenu,
                               QHeaderView, QTextEdit, QGroupBox, QGridLayout, QCheckBox,
                               QSizePolicy, QInputDialog)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QAction, QFont

from src.ui.product_widget import ProductWidget
from src.ui.styles import MAIN_STYLE, SUCCESS_STYLE, WARNING_STYLE, EDIT_STYLE, DELETE_STYLE
from src.models.product import Product
from src.services.excel_service import ExcelService
from src.services.word_service import WordService
from src.services.barcode_generator import BarcodeGenerator
from src.services.file_service import FileService
from src.ui.category_dialog import CategoryDialog
from src.ui.settings_dialog import SettingsDialog

class WorkerThread(QThread):
    """백그라운드 작업 스레드"""
    
    progress_updated = pyqtSignal(int)
    status_updated = pyqtSignal(str)
    finished = pyqtSignal(bool, str)
    

    def __init__(self, excel_service, word_service, products, settings):
        super().__init__()
        self.excel_service = excel_service
        self.word_service = word_service
        self.products = products
        self.settings = settings
        self.data_path = None
    
    def run(self):
        try:
            self.status_updated.emit("바코드 번호 생성 중...")
            self.progress_updated.emit(10)
            
            # Create a new BarcodeGenerator with the options from the settings dialog
            barcode_generator = BarcodeGenerator(self.settings['barcode_options'])
            
            # Word 서비스에 바코드 크기 설정 (셀 크기에 맞춰 최적화)
            template_path = self.settings['template']
            if template_path:
                cell_size = self.word_service.get_cell_size_mm(template_path)
                if cell_size and cell_size[0] > 0 and cell_size[1] > 0:
                    # 셀 크기에 맞춰 바코드 크기 계산 (MM 단위 통일)
                    cell_w_mm, cell_h_mm = cell_size
                    
                    # 바코드 너비: 셀 너비에서 좌우 여백(6mm) 제외
                    barcode_w_mm = max(20.0, min(cell_w_mm - 6.0, 50.0))
                    
                    # 바코드 높이: 셀 높이에서 텍스트 영역(8mm) 제외  
                    barcode_h_mm = max(10.0, min(cell_h_mm - 8.0, 25.0))
                    
                    self.word_service.set_barcode_size_mm(barcode_w_mm, barcode_h_mm)
                    print(f"셀 크기 기반 바코드 크기 설정: {barcode_w_mm:.1f}mm x {barcode_h_mm:.1f}mm")

            items_to_generate = []
            for product in self.products:
                quantity = self.settings['quantities'].get(product.barcode_num, 1)
                for _ in range(quantity):
                    items_to_generate.append(product)

            print(f"생성할 아이템 수: {len(items_to_generate)}")
            items = self.excel_service.generate_barcode_numbers(items_to_generate)
            self.progress_updated.emit(30)
            
            self.status_updated.emit("바코드 이미지 생성 중...")
            barcode_images = barcode_generator.generate_barcodes_for_products(items)
            self.progress_updated.emit(60)
            
            self.status_updated.emit("Word 문서 생성 중...")
            self.word_service.template_file = self.settings['template']
            
            # 단일 파일 생성 여부 확인
            if self.settings.get('single_file', False):
                self.status_updated.emit(f"통합 문서 생성 중... ({len(items)}개 라벨)")
                files_created = self.word_service.generate_single_label_document(items, barcode_images)
            else:
                self.status_updated.emit(f"개별 문서 생성 중... ({len(items)}개 라벨, {len(barcode_images)}개 상품)")
                files_created = self.word_service.generate_label_documents(items, barcode_images)
            
            self.progress_updated.emit(90)
            
            self.status_updated.emit("작업 완료!")
            self.progress_updated.emit(100)
            
            self.finished.emit(True, f"총 {files_created}개 파일이 생성되었습니다.")
            
        except Exception as e:
            self.finished.emit(False, f"오류 발생: {str(e)}")

class MainWindow(QMainWindow):
    """메인 윈도우"""
    
    def __init__(self):
        super().__init__()
        self.products = []
        self.selected_products = []
        self.worker_thread = None
        self.data_path = None
        self.editing_product = None # 상품 수정 시 원본 저장
        
        self.file_service = FileService()
        self.excel_service = None
        self.word_service = None
        
        self.setup_ui()
        self.setup_services()
        self.setup_connections()
        
        self.setStyleSheet(MAIN_STYLE)
    
    def setup_ui(self):
        """UI 구성"""
        self.setWindowTitle("바코드 라벨 생성기")
        self.setGeometry(100, 100, 1200, 800)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout()
        
        title_label = QLabel("바코드 라벨 생성기")
        title_label.setProperty("class", "title")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)
        
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        left_panel = QWidget()
        left_layout = QVBoxLayout()
        
        self.product_widget = ProductWidget()
        left_layout.addWidget(self.product_widget)
        
        file_group = QGroupBox("파일 관리")
        file_layout = QGridLayout()
        
        self.load_excel_button = QPushButton("Excel 파일 불러오기")
        self.load_excel_button.clicked.connect(self.load_excel_file)
        
        self.save_excel_button = QPushButton("Excel 파일 저장")
        self.save_excel_button.clicked.connect(self.save_excel_file)
        
        self.open_output_button = QPushButton("출력 폴더 열기")
        self.open_output_button.clicked.connect(self.open_output_directory)
        
        file_layout.addWidget(self.load_excel_button, 0, 0)
        file_layout.addWidget(self.save_excel_button, 0, 1)
        file_layout.addWidget(self.open_output_button, 1, 0, 1, 2)
        
        file_group.setLayout(file_layout)
        left_layout.addWidget(file_group)
        
        left_panel.setLayout(left_layout)
        splitter.addWidget(left_panel)
        
        right_panel = QWidget()
        right_layout = QVBoxLayout()
        
        table_header_layout = QHBoxLayout()
        table_header_layout.addWidget(QLabel("상품 목록"))
        
        self.select_all_checkbox = QCheckBox("전체 선택")
        table_header_layout.addWidget(self.select_all_checkbox)
        table_header_layout.addStretch()
        right_layout.addLayout(table_header_layout)
        
        self.products_table = QTableWidget()
        self.products_table.setColumnCount(6)
        self.products_table.setHorizontalHeaderLabels(["", "상품명", "가격", "종류", "바코드번호", "관리"])
        self.products_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.products_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.products_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.products_table.setColumnWidth(0, 30)
        self.products_table.setAlternatingRowColors(True)
        self.products_table.setMinimumHeight(400)
        self.products_table.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding)
        self.products_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.products_table.customContextMenuRequested.connect(self.show_table_context_menu)

        right_layout.addWidget(self.products_table)
        
        control_layout = QHBoxLayout()
        
        self.generate_button = QPushButton("라벨 생성 시작")
        self.generate_button.clicked.connect(self.start_generation)
        self.generate_button.setStyleSheet(SUCCESS_STYLE)
        
        self.clear_all_button = QPushButton("전체 삭제")
        self.clear_all_button.clicked.connect(self.clear_all_products)
        self.clear_all_button.setStyleSheet(WARNING_STYLE)
        
        control_layout.addWidget(self.generate_button)
        control_layout.addWidget(self.clear_all_button)
        
        right_layout.addLayout(control_layout)
        
        progress_group = QGroupBox("진행 상황")
        progress_layout = QVBoxLayout()
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        
        self.status_label = QLabel("준비됨")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.status_label)
        
        progress_group.setLayout(progress_layout)
        right_layout.addWidget(progress_group)
        
        log_group = QGroupBox("작업 로그")
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(150)
        self.log_text.setReadOnly(True)
        
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        right_layout.addWidget(log_group)
        
        right_panel.setLayout(right_layout)
        splitter.addWidget(right_panel)
        
        splitter.setSizes([400, 800])
        
        main_layout.addWidget(splitter)
        central_widget.setLayout(main_layout)
        
        self.setup_menu()
    
    def setup_menu(self):
        """메뉴바 설정"""
        menubar = self.menuBar()
        
        file_menu = menubar.addMenu("파일")
        
        load_action = QAction("Excel 파일 불러오기", self)
        load_action.triggered.connect(self.load_excel_file)
        file_menu.addAction(load_action)
        
        save_action = QAction("Excel 파일 저장", self)
        save_action.triggered.connect(self.save_excel_file)
        file_menu.addAction(save_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("종료", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        tools_menu = menubar.addMenu("도구")
        
        manage_categories_action = QAction("종류 관리...", self)
        manage_categories_action.triggered.connect(self.manage_categories)
        tools_menu.addAction(manage_categories_action)
        
        tools_menu.addSeparator()
        
        open_output_action = QAction("출력 폴더 열기", self)
        open_output_action.triggered.connect(self.open_output_directory)
        tools_menu.addAction(open_output_action)
        
        backup_action = QAction("Excel 파일 백업", self)
        backup_action.triggered.connect(self.backup_excel_file)
        tools_menu.addAction(backup_action)
        
        cleanup_action = QAction("임시 파일 정리", self)
        cleanup_action.triggered.connect(self.cleanup_temp_files)
        tools_menu.addAction(cleanup_action)
        
        help_menu = menubar.addMenu("도움말")
        
        about_action = QAction("정보", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
    
    def setup_services(self):
        """서비스 초기화"""
        try:
            self.file_service.ensure_directories()
            
            template_path = self.file_service.get_template_path()
            self.word_service = WordService(template_path)
            
            self.data_path = self.file_service.get_data_path()
            self.excel_service = ExcelService(self.data_path)
            
            self.load_products_from_excel()
            
            self.log_message("서비스 초기화 완료")
            
        except FileNotFoundError as e:
            self.log_message(f"파일을 찾을 수 없습니다: {e}", "error")
            QMessageBox.warning(self, "파일 오류", str(e))
        except Exception as e:
            self.log_message(f"서비스 초기화 실패: {e}", "error")
    
    def load_products_from_excel(self):
        """Excel 파일에서 상품 목록 자동 로드"""
        try:
            categories = self.excel_service.get_categories()
            self.product_widget.set_categories(categories)
            self.log_message(f"종류 목록 로드됨: {list(categories.keys())}")
            
            products = self.excel_service.read_products()
            if products:
                self.products = products
                self.selected_products.clear()
                self.update_products_table()
                self.log_message(f"Excel 파일에서 {len(products)}개 상품을 자동 로드했습니다.")
            else:
                self.log_message("Excel 파일에 상품이 없습니다. 새로 추가해주세요.")
        except Exception as e:
            self.log_message(f"Excel 파일 자동 로드 실패: {e}", "error")
    
    def add_product(self, product: Product):
        """상품 추가 (Excel 파일에 저장)"""
        try:
            if product.type_id == -1: # A new category was entered
                if self.excel_service.add_type_name(product.type_name):
                    self.log_message(f"새 종류 '{product.type_name}'이(가) 추가되었습니다.")
                    categories = self.excel_service.get_categories()
                    self.product_widget.set_categories(categories)
                    if product.type_name in categories:
                        product.type_id = categories[product.type_name]
                    else:
                        self.log_message(f"새 종류 '{product.type_name}'의 ID를 찾을 수 없습니다.", "error")
                        return
                else:
                    self.log_message(f"새 종류 '{product.type_name}' 추가에 실패했습니다.", "error")
                    return

            if self.excel_service.add_product(product):
                self.products = self.excel_service.read_products()
                self.update_products_table()
                self.log_message(f"상품 추가 및 Excel 저장 완료: {product.name}")
            else:
                self.log_message(f"상품 추가 실패: {product.name}", "error")
        except Exception as e:
            self.log_message(f"상품 추가 중 오류: {e}", "error")
    
    def update_product(self, updated_product: Product):
        """상품 수정 (Excel 파일에 저장)"""
        try:
            old_product = self.editing_product
            if not old_product:
                self.log_message("수정할 원본 상품 정보를 찾을 수 없습니다.", "error")
                return

            if updated_product.type_id == -1: # A new category was entered
                if self.excel_service.add_type_name(updated_product.type_name):
                    self.log_message(f"새 종류 '{updated_product.type_name}'이(가) 추가되었습니다.")
                    categories = self.excel_service.get_categories()
                    self.product_widget.set_categories(categories)
                    if updated_product.type_name in categories:
                        updated_product.type_id = categories[updated_product.type_name]
                    else:
                        self.log_message(f"새 종류 '{updated_product.type_name}'의 ID를 찾을 수 없습니다.", "error")
                        return
                else:
                    self.log_message(f"새 종류 '{updated_product.type_name}' 추가에 실패했습니다.", "error")
                    return

            if self.excel_service.update_product(old_product, updated_product):
                self.products = self.excel_service.read_products()
                self.update_products_table()
                self.log_message(f"상품 수정 및 Excel 저장 완료: {updated_product.name}")
            else:
                self.log_message(f"상품 수정 실패: {updated_product.name}", "error")
            
            self.editing_product = None

        except Exception as e:
            self.log_message(f"상품 수정 중 오류: {e}", "error")
            self.editing_product = None

    def delete_product(self, product: Product):
        """상품 삭제 (Excel 파일에서도 삭제)"""
        reply = QMessageBox.question(self, "상품 삭제", 
                                   f"'{product.name}' 상품을 삭제하시겠습니까?\n(Excel 파일에서도 삭제됩니다)",
                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            try:
                if self.excel_service.delete_product(product):
                    self.products.remove(product)
                    if product in self.selected_products:
                        self.selected_products.remove(product)
                    self.update_products_table()
                    self.log_message(f"상품 삭제 및 Excel 저장 완료: {product.name}")
                else:
                    self.log_message(f"상품 삭제 실패: {product.name}", "error")
            except Exception as e:
                self.log_message(f"상품 삭제 중 오류: {e}", "error")
    
    def clear_all_products(self):
        """모든 상품 삭제 (Excel 파일도 초기화)"""
        if not self.products:
            return
        
        reply = QMessageBox.question(self, "전체 삭제", 
                                   f"모든 상품({len(self.products)}개)을 삭제하시겠습니까?\n(Excel 파일도 초기화됩니다)",
                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            try:
                if self.excel_service.save_products([], self.data_path):
                    self.products.clear()
                    self.selected_products.clear()
                    self.update_products_table()
                    self.log_message("모든 상품 삭제 및 Excel 초기화 완료")
                else:
                    self.log_message("Excel 파일 초기화 실패", "error")
            except Exception as e:
                self.log_message(f"전체 삭제 중 오류: {e}", "error")
    
    def load_excel_file(self):
        """Excel 파일 불러오기 (다른 파일 선택)"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Excel 파일 선택", "", "Excel Files (*.xlsx *.xls)")
        
        if file_path:
            try:
                temp_excel_service = ExcelService(file_path)
                
                categories = temp_excel_service.get_categories()
                self.product_widget.set_categories(categories)
                
                products = temp_excel_service.read_products()
                
                if products or categories:
                    self.excel_service = temp_excel_service
                    self.products = products
                    self.selected_products.clear()
                    self.update_products_table()
                    self.log_message(f"새 Excel 파일에서 {len(products)}개 상품, {len(categories)}개 종류를 불러왔습니다.")
                else:
                    self.log_message("선택한 Excel 파일에서 데이터를 찾을 수 없습니다.", "warning")
                    
            except Exception as e:
                self.log_message(f"Excel 파일 로드 실패: {e}", "error")
                QMessageBox.critical(self, "오류", f"Excel 파일 로드 실패: {e}")
    
    def save_excel_file(self):
        """Excel 파일 저장 (현재 데이터를 다른 위치에 저장)"""
        if not self.products:
            QMessageBox.information(self, "알림", "저장할 상품이 없습니다.")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Excel 파일 저장", "products.xlsx", "Excel Files (*.xlsx)")
        
        if file_path:
            try:
                if self.excel_service.save_products(self.products, file_path):
                    self.log_message(f"Excel 파일 저장 완료: {file_path}")
                    QMessageBox.information(self, "완료", "Excel 파일이 저장되었습니다.")
                else:
                    self.log_message("Excel 파일 저장 실패", "error")
                    QMessageBox.critical(self, "오류", "Excel 파일 저장에 실패했습니다.")
                    
            except Exception as e:
                self.log_message(f"Excel 파일 저장 실패: {e}", "error")
                QMessageBox.critical(self, "오류", f"Excel 파일 저장 실패: {e}")
    
    def backup_excel_file(self):
        """Excel 파일 백업"""
        try:
            if self.excel_service.backup_file():
                self.log_message("Excel 파일 백업 완료")
                QMessageBox.information(self, "완료", "Excel 파일이 백업되었습니다.")
            else:
                self.log_message("Excel 파일 백업 실패", "error")
                QMessageBox.critical(self, "오류", "Excel 파일 백업에 실패했습니다.")
        except Exception as e:
            self.log_message(f"Excel 파일 백업 중 오류: {e}", "error")
    
    def start_generation(self):
        """라벨 생성 시작"""
        if not self.selected_products:
            QMessageBox.warning(self, "경고", "생성할 상품을 선택해주세요.")
            return
        
        templates = [self.file_service.get_template_path(f) for f in os.listdir(self.file_service.get_template_directory()) if f.endswith('.docx')]

        max_table_size = [self.word_service.get_table_max_size(template) for template in templates]
        cell_sizes = [self.word_service.get_cell_size_mm(template) for template in templates]
        dialog = SettingsDialog(templates, self.selected_products, self, max_table_size, cell_sizes)

        if dialog.exec():
            settings = dialog.get_settings()
            self.generate_labels(settings)
    
    def generate_labels(self, settings):
        """라벨 생성 실행"""
        try:
            self.generate_button.setEnabled(False)
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            
            self.worker_thread = WorkerThread(
                self.excel_service, self.word_service, self.selected_products, settings
            )
            
            self.worker_thread.progress_updated.connect(self.progress_bar.setValue)
            self.worker_thread.status_updated.connect(self.status_label.setText)
            self.worker_thread.finished.connect(self.generation_finished)
            
            self.worker_thread.start()
            
        except Exception as e:
            self.log_message(f"라벨 생성 시작 실패: {e}", "error")
            self.generation_finished(False, str(e))
    
    def generation_finished(self, success: bool, message: str):
        """라벨 생성 완료"""
        self.generate_button.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        if success:
            self.status_label.setText("완료!")
            self.log_message(message, "success")
            QMessageBox.information(self, "완료", message)
        else:
            self.status_label.setText("실패")
            self.log_message(message, "error")
            QMessageBox.critical(self, "오류", message)
    
    def open_output_directory(self):
        """출력 폴더 열기"""
        try:
            self.file_service.open_output_directory()
        except Exception as e:
            self.log_message(f"출력 폴더 열기 실패: {e}", "error")
    
    def cleanup_temp_files(self):
        """임시 파일 정리"""
        reply = QMessageBox.question(self, "파일 정리", 
                                   "출력 파일을 정리하시겠습니까?",
                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            try:
                self.file_service.cleanup_output()
                self.log_message("임시 파일 정리 완료")
                QMessageBox.information(self, "완료", "출력 파일이 정리되었습니다.")
            except Exception as e:
                self.log_message(f"파일 정리 실패: {e}", "error")
    
    def show_about(self):
        """정보 표시"""
        QMessageBox.about(self, "바코드 라벨 생성기", 
                         "바코드 라벨 생성기 v1.0\n\n"
                         "상품 정보를 입력하여 바코드가 포함된 라벨을 생성합니다.\n"
                         "복사 옵션을 선택하면 한 페이지에 78개의 라벨을 생성합니다.")
    
    def log_message(self, message: str, level: str = "info"):
        """로그 메시지 추가"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        if level == "error":
            prefix = "[오류]"
            color = "#d83b01"
        elif level == "warning":
            prefix = "[경고]"
            color = "#d83b01"
        elif level == "success":
            prefix = "[성공]"
            color = "#107c10"
        else:
            prefix = "[정보]"
            color = "#666666"
        
        log_entry = f'<span style="color: {color};">{timestamp} {prefix}</span> {message}'
        self.log_text.append(log_entry)
        
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
    
    def closeEvent(self, event):
        """프로그램 종료 시 처리"""
        if self.worker_thread and self.worker_thread.isRunning():
            reply = QMessageBox.question(self, "작업 중", 
                                       "라벨 생성이 진행 중입니다. 정말 종료하시겠습니까?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            
            if reply == QMessageBox.StandardButton.Yes:
                self.worker_thread.terminate()
                self.worker_thread.wait()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

    def setup_connections(self):
        """시그널 연결"""
        self.product_widget.productAdded.connect(self.add_product)
        self.product_widget.productUpdated.connect(self.update_product)
        self.select_all_checkbox.stateChanged.connect(self._on_select_all_changed)
    
    def update_products_table(self):
        """상품 테이블 업데이트"""
        self.products_table.setRowCount(len(self.products))
        
        for row, product in enumerate(self.products):
            checkbox = QCheckBox()
            checkbox.setChecked(product in self.selected_products)
            checkbox.stateChanged.connect(lambda state, p=product: self._on_product_selected(state, p))
            
            cell_widget = QWidget()
            layout = QHBoxLayout(cell_widget)
            layout.addWidget(checkbox)
            layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.setContentsMargins(0, 0, 0, 0)
            self.products_table.setCellWidget(row, 0, cell_widget)
            
            name_item = QTableWidgetItem(product.name)
            # make non-editable
            name_item.setFlags(name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.products_table.setItem(row, 1, name_item)
            
            price_item = QTableWidgetItem(product.formatted_price)
            price_item.setFlags(price_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.products_table.setItem(row, 2, price_item)
            
            category_item = QTableWidgetItem(product.type_name)
            category_item.setFlags(category_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.products_table.setItem(row, 3, category_item)
            
            item_number = str(product.product_id).zfill(6)
            barcode_item = QTableWidgetItem(f"{product.type_id}{item_number}" if product.type_id is not None and product.product_id is not None else "")
            barcode_item.setFlags(barcode_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.products_table.setItem(row, 4, barcode_item)
            
            button_widget = QWidget()
            button_layout = QHBoxLayout()
            button_layout.setContentsMargins(2, 2, 2, 2)
            
            edit_button = QPushButton("✏️ 수정")
            edit_button.clicked.connect(lambda checked, p=product: self.edit_product(p))
            edit_button.setMaximumWidth(70)
            edit_button.setStyleSheet(EDIT_STYLE)
            
            delete_button = QPushButton("🗑️ 삭제")
            delete_button.clicked.connect(lambda checked, p=product: self.delete_product(p))
            delete_button.setMaximumWidth(70)
            delete_button.setStyleSheet(DELETE_STYLE)
            
            button_layout.addWidget(edit_button)
            button_layout.addWidget(delete_button)
            button_widget.setLayout(button_layout)
            
            self.products_table.setCellWidget(row, 5, button_widget)
        
        self._update_select_all_checkbox_state()

    def _on_select_all_changed(self, state):
        """'전체 선택' 체크박스 상태 변경 시 호출"""
        if self.select_all_checkbox.isTristate():
            self.select_all_checkbox.setTristate(False)
        
        if state == Qt.CheckState.Checked.value:
            self.selected_products = list(self.products)
        else:
            self.selected_products.clear()
        self.update_products_table()

    def _on_product_selected(self, state, product):
        """개별 상품 체크박스 상태 변경 시 호출"""
        if state == Qt.CheckState.Checked.value:
            if product not in self.selected_products:
                self.selected_products.append(product)
        else:
            if product in self.selected_products:
                self.selected_products.remove(product)
        self._update_select_all_checkbox_state()

    def _update_select_all_checkbox_state(self):
        """'전체 선택' 체크박스 상태를 동기화"""
        self.select_all_checkbox.blockSignals(True)
        
        if not self.products:
            self.select_all_checkbox.setTristate(False)
            self.select_all_checkbox.setCheckState(Qt.CheckState.Unchecked)
        elif len(self.selected_products) == len(self.products):
            self.select_all_checkbox.setTristate(False)
            self.select_all_checkbox.setCheckState(Qt.CheckState.Checked)
        elif not self.selected_products:
            self.select_all_checkbox.setTristate(False)
            self.select_all_checkbox.setCheckState(Qt.CheckState.Unchecked)
        else:
            self.select_all_checkbox.setTristate(True)
            self.select_all_checkbox.setCheckState(Qt.CheckState.PartiallyChecked)
            
        self.select_all_checkbox.blockSignals(False)

    def edit_product(self, product: Product):
        """상품 수정 모드로 전환"""
        self.editing_product = product
        self.product_widget.edit_product(product)

    def manage_categories(self):
        """종류 관리 다이얼로그 열기"""
        if not self.excel_service:
            QMessageBox.warning(self, "오류", "Excel 서비스가 초기화되지 않았습니다.")
            return
        
        dialog = CategoryDialog(self.excel_service, self)
        dialog.categories_updated.connect(self._on_categories_updated)
        dialog.exec()

    def _on_categories_updated(self):
        """종류 변경 시 UI 업데이트"""
        self.log_message("종류 목록이 변경되어 데이터를 새로고침합니다.")
        
        categories = self.excel_service.get_categories()
        self.product_widget.set_categories(categories)
        
        self.products = self.excel_service.read_products()
        self.selected_products.clear()
        
        self.update_products_table()
        
        self.product_widget.clear_form()

    def show_table_context_menu(self, pos):
        """상품 테이블 컨텍스트 메뉴 표시"""
        item = self.products_table.itemAt(pos)
        if not item:
            return

        row = item.row()
        product = self.products[row]

        menu = QMenu()
        change_id_action = menu.addAction("제품ID 변경...")
        action = menu.exec(self.products_table.mapToGlobal(pos))

        if action == change_id_action:
            self.change_product_id(product)

    def change_product_id(self, product: Product):
        """제품ID 변경 (입력 다이얼로그) — 중복 시 교환 확인"""
        try:
            current = int(getattr(product, "product_id", 0) or 0)
        except Exception:
            current = 0

        new_id, ok = QInputDialog.getInt(self, "제품ID 변경", f"'{product.name}'의 새로운 제품ID를 입력하세요:", value=current, min=0)
        if not ok:
            return

        if new_id == current:
            return

        existing = next((p for p in self.products if int(getattr(p, "product_id", 0) or 0) == new_id), None)

        if existing and existing is not product:
            resp = QMessageBox.question(
                self, 
                "제품ID 중복",
                f"제품ID {new_id}은(는) 이미 '{existing.name}'에서 사용 중입니다.\n"
                f"두 상품의 제품ID를 서로 교환하시겠습니까?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if resp == QMessageBox.StandardButton.Yes:
                try:
                    existing_id = int(getattr(existing, "product_id", 0) or 0)
                    existing.product_id = current
                    product.product_id = new_id

                    if self.excel_service and self.excel_service.save_products(self.products, self.data_path):
                        self.update_products_table()
                        self.log_message(f"제품ID 교환 완료: {product.name}({current}→{new_id}) <-> {existing.name}({existing_id}→{current})", "success")
                        QMessageBox.information(self, "완료", "제품ID가 교환되어 저장되었습니다.")
                    else:
                        existing.product_id = existing_id
                        product.product_id = current
                        QMessageBox.critical(self, "오류", "제품ID 저장에 실패했습니다.")
                        self.log_message("제품ID 저장 실패", "error")
                except Exception as e:
                    existing.product_id = int(getattr(existing, "product_id", 0) or 0)
                    product.product_id = current
                    QMessageBox.critical(self, "오류", f"제품ID 교환 중 오류: {e}")
                    self.log_message(f"제품ID 교환 오류: {e}", "error")
            else:
                return
        else:
            old_id = current
            product.product_id = new_id
            try:
                if self.excel_service and self.excel_service.save_products(self.products, self.data_path):
                    self.update_products_table()
                    self.log_message(f"제품ID 변경 완료: {product.name} ({old_id} → {new_id})", "success")
                    QMessageBox.information(self, "완료", "제품ID가 변경되어 저장되었습니다.")
                else:
                    product.product_id = old_id
                    QMessageBox.critical(self, "오류", "제품ID 저장에 실패했습니다.")
                    self.log_message("제품ID 저장 실패", "error")
            except Exception as e:
                product.product_id = old_id
                QMessageBox.critical(self, "오류", f"제품ID 변경 실패: {e}")
                self.log_message(f"제품ID 변경 중 오류: {e}", "error")
