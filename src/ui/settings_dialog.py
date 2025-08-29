from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QComboBox, QSpinBox, QDoubleSpinBox,
    QDialogButtonBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QLabel, QGroupBox
)
from PyQt6.QtCore import Qt
import os

class SettingsDialog(QDialog):
    """라벨 생성 세부 설정 다이얼로그"""
    def __init__(self, templates, products, parent=None):
        super().__init__(parent)
        self.setWindowTitle("라벨 생성 세부 설정")
        self.setMinimumSize(500, 600)

        self.templates = templates
        self.products = products

        layout = QVBoxLayout(self)

        # 1. 템플릿 선택
        template_group = QGroupBox("템플릿 설정")
        template_layout = QFormLayout()
        self.template_combo = QComboBox()
        self.template_combo.addItems([os.path.basename(t) for t in self.templates])
        template_layout.addRow("라벨 양식 선택:", self.template_combo)
        template_group.setLayout(template_layout)
        layout.addWidget(template_group)

        # 2. 바코드 설정
        barcode_group = QGroupBox("바코드 설정")
        barcode_layout = QFormLayout()
        self.module_width_spin = QDoubleSpinBox()
        self.module_width_spin.setDecimals(2)
        self.module_width_spin.setSingleStep(0.1)
        self.module_width_spin.setValue(0.2)
        barcode_layout.addRow("모듈 너비 (mm):", self.module_width_spin)

        self.module_height_spin = QDoubleSpinBox()
        self.module_height_spin.setDecimals(1)
        self.module_height_spin.setSingleStep(1.0)
        self.module_height_spin.setValue(15.0)
        barcode_layout.addRow("모듈 높이 (mm):", self.module_height_spin)

        self.quiet_zone_spin = QDoubleSpinBox()
        self.quiet_zone_spin.setDecimals(1)
        self.quiet_zone_spin.setSingleStep(0.5)
        self.quiet_zone_spin.setValue(6.5)
        barcode_layout.addRow("좌우 여백 (mm):", self.quiet_zone_spin)

        self.font_size_spin = QSpinBox()
        self.font_size_spin.setValue(10)
        barcode_layout.addRow("글자 크기 (pt):", self.font_size_spin)

        self.dpi_spin = QSpinBox()
        self.dpi_spin.setRange(100, 600)
        self.dpi_spin.setSingleStep(50)
        self.dpi_spin.setValue(300)
        barcode_layout.addRow("DPI:", self.dpi_spin)
        
        barcode_group.setLayout(barcode_layout)
        layout.addWidget(barcode_group)

        # 3. 상품별 출력 개수 설정
        product_group = QGroupBox("상품별 출력 개수")
        product_layout = QVBoxLayout()
        self.product_table = QTableWidget()
        self.product_table.setColumnCount(2)
        self.product_table.setHorizontalHeaderLabels(["상품명", "출력 개수"])
        self.product_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.product_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.product_table.setRowCount(len(self.products))

        for i, product in enumerate(self.products):
            name_item = QTableWidgetItem(product.name)
            name_item.setFlags(name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.product_table.setItem(i, 0, name_item)

            spin_box = QSpinBox()
            spin_box.setMinimum(1)
            spin_box.setMaximum(999)
            spin_box.setValue(1)
            self.product_table.setCellWidget(i, 1, spin_box)

        product_layout.addWidget(self.product_table)
        product_group.setLayout(product_layout)
        layout.addWidget(product_group)

        # 확인/취소 버튼
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

    def get_settings(self):
        """설정된 값들을 반환하는 메서드"""
        
        # 선택된 템플릿 경로 찾기
        selected_template_name = self.template_combo.currentText()
        selected_template_path = next((t for t in self.templates if os.path.basename(t) == selected_template_name), None)

        # 바코드 설정
        barcode_options = {
            'module_width': self.module_width_spin.value(),
            'module_height': self.module_height_spin.value(),
            'quiet_zone': self.quiet_zone_spin.value(),
            'font_size': self.font_size_spin.value(),
            'dpi': self.dpi_spin.value(),
        }

        # 상품별 출력 개수
        quantities = {}
        for i, product in enumerate(self.products):
            spin_box = self.product_table.cellWidget(i, 1)
            quantities[product.product_id] = spin_box.value()

        return {
            "template": selected_template_path,
            "quantities": quantities,
            "barcode_options": barcode_options
        }
