import os
import sys
from PyQt6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                               QSplitter, QTableWidget, QTableWidgetItem, QPushButton,
                               QLabel, QProgressBar, QMessageBox, QFileDialog, QMenu,
                               QHeaderView, QTextEdit, QGroupBox, QGridLayout, QCheckBox,
                               QSizePolicy)
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

class WorkerThread(QThread):
    """백그라운드 작업 스레드"""
    
    progress_updated = pyqtSignal(int)
    status_updated = pyqtSignal(str)
    finished = pyqtSignal(bool, str)
    

    def __init__(self, excel_service, barcode_generator, word_service, products):
        super().__init__()
        self.excel_service = excel_service
        self.barcode_generator = barcode_generator
        self.word_service = word_service
        self.products = products
        self.data_path = None
    
    def run(self):
        try:
            self.status_updated.emit("바코드 번호 생성 중...")
            self.progress_updated.emit(10)
            
            # 바코드 번호 생성
            items = self.excel_service.generate_barcode_numbers(self.products)
            self.progress_updated.emit(30)
            
            self.status_updated.emit("바코드 이미지 생성 중...")
            # 바코드 이미지를 메모리에 생성
            barcode_images = self.barcode_generator.generate_barcodes_for_products(items)
            self.progress_updated.emit(60)
            
            self.status_updated.emit("Word 문서 생성 중...")
            # Word 문서 생성 (메모리의 바코드 이미지 사용)
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
        
        # 서비스 초기화
        self.file_service = FileService()
        self.excel_service = None
        self.word_service = None
        self.barcode_generator = None
        
        self.setup_ui()
        self.setup_services()
        self.setup_connections()
        
        # 스타일 적용
        self.setStyleSheet(MAIN_STYLE)
    
    def setup_ui(self):
        """UI 구성"""
        self.setWindowTitle("바코드 라벨 생성기")
        self.setGeometry(100, 100, 1200, 800)
        
        # 중앙 위젯
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 메인 레이아웃
        main_layout = QVBoxLayout()
        
        # 제목
        title_label = QLabel("바코드 라벨 생성기")
        title_label.setProperty("class", "title")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)
        
        # 스플리터로 좌우 분할
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # 왼쪽 패널 (상품 입력)
        left_panel = QWidget()
        left_layout = QVBoxLayout()
        
        self.product_widget = ProductWidget()
        left_layout.addWidget(self.product_widget)
        
        # 파일 관련 버튼들
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
        
        # 오른쪽 패널 (상품 목록 및 제어)
        right_panel = QWidget()
        right_layout = QVBoxLayout()
        
        # 상품 목록 헤더 (라벨 + 전체 선택 체크박스)
        table_header_layout = QHBoxLayout()
        table_header_layout.addWidget(QLabel("상품 목록"))
        
        self.select_all_checkbox = QCheckBox("전체 선택")
        table_header_layout.addWidget(self.select_all_checkbox)
        table_header_layout.addStretch()
        right_layout.addLayout(table_header_layout)
        
        # 상품 목록 테이블
        self.products_table = QTableWidget()
        self.products_table.setColumnCount(6)
        self.products_table.setHorizontalHeaderLabels(["", "상품명", "가격", "종류", "복사", "관리"])
        self.products_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.products_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.products_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.products_table.setColumnWidth(0, 30)
        self.products_table.setAlternatingRowColors(True)
        self.products_table.setMinimumHeight(400)  # 최소 높이 설정
        self.products_table.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding)  # 세로 방향으로 확장
        
        right_layout.addWidget(self.products_table)
        
        # 제어 버튼들
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
        
        # 진행 상황
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
        
        # 로그 출력
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
        
        # 스플리터 비율 설정
        splitter.setSizes([400, 800])
        
        main_layout.addWidget(splitter)
        central_widget.setLayout(main_layout)
        
        # 메뉴바
        self.setup_menu()
    
    def setup_menu(self):
        """메뉴바 설정"""
        menubar = self.menuBar()
        
        # 파일 메뉴
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
        
        # 도구 메뉴
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
        
        # 도움말 메뉴
        help_menu = menubar.addMenu("도움말")
        
        about_action = QAction("정보", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
    
    def setup_services(self):
        """서비스 초기화"""
        try:
            self.file_service.ensure_directories()
            
            # 템플릿 파일 경로 가져오기
            template_path = self.file_service.get_template_path()
            self.word_service = WordService(template_path)
            
            # 바코드 생성기 초기화
            self.barcode_generator = BarcodeGenerator()
            
            # Excel 서비스 초기화
            self.data_path = self.file_service.get_data_path()
            self.excel_service = ExcelService(self.data_path)
            
            # 프로그램 시작 시 Excel 파일에서 상품 목록 자동 로드
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
            # 종류 목록을 ProductWidget에 설정
            categories = self.excel_service.get_categories()
            self.product_widget.set_categories(categories)
            self.log_message(f"종류 목록 로드됨: {categories}")
            
            # 상품 목록 로드
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
            # 새로운 종류인 경우 type 시트에 추가
            if product.category not in self.excel_service.get_categories():
                if self.excel_service.add_category(product.category):
                    # ProductWidget의 종류 목록 업데이트
                    categories = self.excel_service.get_categories()
                    self.product_widget.set_categories(categories)
                    self.log_message(f"새 종류 추가됨: {product.category}")
            
            # Excel 파일에 상품 추가
            if self.excel_service.add_product(product):
                self.products.append(product)
                self.update_products_table()
                self.log_message(f"상품 추가 및 Excel 저장 완료: {product.name}")
            else:
                self.log_message(f"상품 추가 실패: {product.name}", "error")
        except Exception as e:
            self.log_message(f"상품 추가 중 오류: {e}", "error")
    
    def update_product(self, product: Product):
        """상품 수정 (Excel 파일에 저장)"""
        try:
            # 기존 상품 찾아서 교체
            old_product = None
            for i, existing_product in enumerate(self.products):
                if (
                    existing_product.name == product.name and 
                    existing_product.category == product.category
                ):
                    old_product = existing_product
                    self.products[i] = product
                    break
            
            if old_product:
                # Excel 파일에 상품 수정
                if self.excel_service.update_product(old_product, product):
                    self.update_products_table()
                    self.log_message(f"상품 수정 및 Excel 저장 완료: {product.name}")
                else:
                    self.log_message(f"상품 수정 실패: {product.name}", "error")
            else:
                self.log_message(f"수정할 상품을 찾을 수 없습니다: {product.name}", "error")
                
        except Exception as e:
            self.log_message(f"상품 수정 중 오류: {e}", "error")
    
    def delete_product(self, product: Product):
        """상품 삭제 (Excel 파일에서도 삭제)"""
        reply = QMessageBox.question(self, "상품 삭제", 
                                   f"'{product.name}' 상품을 삭제하시겠습니까?\n(Excel 파일에서도 삭제됩니다)",
                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            try:
                # Excel 파일에서 상품 삭제
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
                # Excel 파일 초기화
                if self.excel_service.save_products([],self.data_path):
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
                # 새로운 Excel 서비스 생성
                temp_excel_service = ExcelService(file_path)
                
                # 종류 목록 업데이트
                categories = temp_excel_service.get_categories()
                self.product_widget.set_categories(categories)
                
                # 상품 목록 로드
                products = temp_excel_service.read_products()
                
                if products or categories:
                    # 현재 Excel 서비스 교체
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
                # 현재 상품 목록을 선택한 위치에 저장
                # temp_excel_service = ExcelService(file_path)
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
        
        reply = QMessageBox.question(self, "라벨 생성", 
                                   f"{len(self.selected_products)}개 상품에 대해 라벨을 생성하시겠습니까?",
                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            self.generate_labels()
    
    def generate_labels(self):
        """라벨 생성 실행"""
        try:
            # UI 비활성화
            self.generate_button.setEnabled(False)
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            
            # 작업 스레드 시작
            self.worker_thread = WorkerThread(
                self.excel_service, self.barcode_generator, self.word_service, self.selected_products
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
        # UI 복원
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
        
        # 자동 스크롤
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
            # 선택 체크박스
            checkbox = QCheckBox()
            checkbox.setChecked(product in self.selected_products)
            checkbox.stateChanged.connect(lambda state, p=product: self._on_product_selected(state, p))
            
            cell_widget = QWidget()
            layout = QHBoxLayout(cell_widget)
            layout.addWidget(checkbox)
            layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.setContentsMargins(0, 0, 0, 0)
            self.products_table.setCellWidget(row, 0, cell_widget)
            
            # 상품명
            name_item = QTableWidgetItem(product.name)
            self.products_table.setItem(row, 1, name_item)
            
            # 가격
            price_item = QTableWidgetItem(product.formatted_price)
            self.products_table.setItem(row, 2, price_item)
            
            # 종류
            category_item = QTableWidgetItem(product.category)
            self.products_table.setItem(row, 3, category_item)
            
            # 복사
            copy_item = QTableWidgetItem("예" if product.copy else "아니오")
            self.products_table.setItem(row, 4, copy_item)
            
            # 관리 버튼
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
        
        # 1. ProductWidget의 종류 목록 업데이트
        categories = self.excel_service.get_categories()
        self.product_widget.set_categories(categories)
        
        # 2. 현재 상품 목록 다시 로드 (변경된 종류 이름 반영)
        self.products = self.excel_service.read_products()
        self.selected_products.clear()
        
        # 3. 상품 테이블 업데이트
        self.update_products_table()
        
        # 4. 수정 중이던 상품 폼 초기화
        self.product_widget.clear_form()