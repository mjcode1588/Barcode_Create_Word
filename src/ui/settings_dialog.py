from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QComboBox, QSpinBox, QDoubleSpinBox,
    QDialogButtonBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QLabel, QGroupBox,QCheckBox
)
from PyQt6.QtCore import Qt
import os

class SettingsDialog(QDialog):
    """라벨 생성 세부 설정 다이얼로그"""
    def __init__(self, templates, products, parent=None, template_area=None, cell_sizes=None):
        super().__init__(parent)
        self.setWindowTitle("라벨 생성 세부 설정")
        self.setMinimumSize(500, 600)

        self.templates = templates
        self.products = products
        self.template_table_size_list = template_area
        self.cell_sizes = cell_sizes

        # 저장해 둘 체크박스/스핀박스 참조
        self.max_checkboxes = {}

        layout = QVBoxLayout(self)

        # 1. 템플릿 선택
        template_group = QGroupBox("템플릿 설정")
        template_layout = QFormLayout()
        self.template_combo = QComboBox()
        self.template_combo.addItems([os.path.basename(t) for t in self.templates])
        template_layout.addRow("라벨 양식 선택:", self.template_combo)
        template_group.setLayout(template_layout)
        layout.addWidget(template_group)

        # 템플릿별 최대 라벨 수 표시 라벨
        self.template_max_label = QLabel("-")
        template_layout.addRow("템플릿 최대 라벨 수:", self.template_max_label)

        template_group.setLayout(template_layout)
        layout.addWidget(template_group)

        # 템플릿 변경 시 처리
        self.template_combo.currentIndexChanged.connect(self.on_template_changed)


        # 2. 바코드 설정 (MM 단위 통일, 템플릿 변경 시 자동 최적화)
        barcode_group = QGroupBox("바코드 설정 (단위: MM)")
        barcode_layout = QFormLayout()
        
        # 설명 라벨 추가
        info_label = QLabel("※ 템플릿 변경 시 셀 크기에 맞춰 자동으로 최적화됩니다")
        info_label.setStyleSheet("color: #666; font-size: 11px;")
        barcode_layout.addRow(info_label)
        
        self.module_width_spin = QDoubleSpinBox()
        self.module_width_spin.setDecimals(2)
        self.module_width_spin.setRange(0.2, 1.0)  # 최소값을 0.2mm로 증가
        self.module_width_spin.setSingleStep(0.05)
        self.module_width_spin.setValue(0.3)  # 기본값을 0.3mm로 증가
        self.module_width_spin.setToolTip("바코드 한 모듈의 너비 (0.2-1.0mm)")
        barcode_layout.addRow("모듈 너비 (mm):", self.module_width_spin)

        self.module_height_spin = QDoubleSpinBox()
        self.module_height_spin.setDecimals(1)
        self.module_height_spin.setRange(5.0, 25.0)
        self.module_height_spin.setSingleStep(1.0)
        self.module_height_spin.setValue(15.0)
        self.module_height_spin.setToolTip("바코드 높이 (5.0-25.0mm)")
        barcode_layout.addRow("바코드 높이 (mm):", self.module_height_spin)

        self.quiet_zone_spin = QDoubleSpinBox()
        self.quiet_zone_spin.setDecimals(1)
        self.quiet_zone_spin.setRange(2.5, 5.0)  # 최소값을 2.5mm로 증가
        self.quiet_zone_spin.setSingleStep(0.5)
        self.quiet_zone_spin.setValue(3.0)  # 기본값을 3.0mm로 증가
        self.quiet_zone_spin.setToolTip("바코드 좌우 여백 (2.5-5.0mm)")
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

        # 3. 출력 설정
        output_group = QGroupBox("출력 설정")
        output_layout = QFormLayout()
        
        # 파일 생성 방식 선택
        self.single_file_checkbox = QCheckBox("하나의 파일로 생성")
        self.single_file_checkbox.setToolTip("체크하면 모든 라벨을 하나의 DOCX 파일로 합쳐서 생성합니다.\n" +
                                           "체크하지 않으면 상품별로 개별 파일을 생성합니다.")
        self.single_file_checkbox.setChecked(False)
        
        # 단일 파일 생성 시 추가 정보 표시
        self.single_file_checkbox.stateChanged.connect(self._on_single_file_changed)
        
        output_layout.addRow("파일 생성 방식:", self.single_file_checkbox)
        
        # 단일 파일 생성 시 안내 라벨
        self.single_file_info = QLabel("※ 하나의 파일로 생성 시 '통합_라벨.docx' 파일이 생성됩니다")
        self.single_file_info.setStyleSheet("color: #666; font-size: 11px;")
        self.single_file_info.setVisible(False)
        output_layout.addRow(self.single_file_info)
        
        output_group.setLayout(output_layout)
        layout.addWidget(output_group)

        # 4. 상품별 출력 개수 설정
        product_group = QGroupBox("상품별 출력 개수")
        product_layout = QVBoxLayout()
        self.product_table = QTableWidget()
        # 컬럼: 상품명 | 출력 개수 | Max 체크박스
        self.product_table.setColumnCount(3)
        self.product_table.setHorizontalHeaderLabels(["상품명", "출력 개수", "Max"])
        self.product_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.product_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.product_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.product_table.setRowCount(len(self.products))
        self.product_table.verticalHeader().setDefaultSectionSize(36)

        for i, product in enumerate(self.products):
            name_item = QTableWidgetItem(product.name)
            name_item.setFlags(name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.product_table.setItem(i, 0, name_item)

            spin_box = QSpinBox()
            spin_box.setMinimum(1)
            spin_box.setMaximum(999)
            spin_box.setValue(1)
            self.product_table.setCellWidget(i, 1, spin_box)

            # Max 체크박스
            chk = QCheckBox("Max")
            chk.setToolTip("체크하면 템플릿의 최대 라벨 수로 설정됩니다.")
            # 연결: row 인덱스 보존하려면 기본값 인자로 캡처
            chk.stateChanged.connect(lambda state, row=i: self.on_max_checked(row, state))
            self.product_table.setCellWidget(i, 2, chk)
            self.max_checkboxes[i] = chk


        product_layout.addWidget(self.product_table)
        product_group.setLayout(product_layout)
        layout.addWidget(product_group)

        # 확인/취소 버튼
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)
        
        # 초기 설정
        self.update_max_label_and_checkboxes()
        # 첫 번째 템플릿에 대해 바코드 크기 자동 최적화
        if self.templates:
            self.template_auto_size(0)

    def get_current_template_max(self):
        """현재 선택된 템플릿에 대한 최대 라벨 수를 반환하거나 None."""
        idx = self.template_combo.currentIndex()
        if self.template_table_size_list is None:
            return None

        # dict로 전달된 경우: 키가 전체 경로 또는 basename 일 수 있음
        if isinstance(self.template_table_size_list, dict):
            # 우선 전체 경로로 시도
            selected_path = self.templates[idx] if 0 <= idx < len(self.templates) else None
            if selected_path and selected_path in self.template_table_size_list:
                return self.template_table_size_list[selected_path]
            # basename으로도 시도
            base = os.path.basename(selected_path) if selected_path else None
            if base and base in self.template_table_size_list:
                return self.template_table_size_list[base]
            # 마지막으로 index 기반으로 찾기 (값이 리스트처럼 들어있을 때)
            return None

        # list/tuple 로 전달된 경우: 길이가 templates와 같으면 인덱스로 반환
        if isinstance(self.template_table_size_list, (list, tuple)):
            if 0 <= idx < len(self.template_table_size_list):
                return self.template_table_size_list[idx]
            return None

        # 단일 숫자인 경우
        try:
            return int(self.template_table_size_list)
        except Exception:
            return None

    def update_max_label_and_checkboxes(self):
        max_val = self.get_current_template_max()
        if max_val is None:
            self.template_max_label.setText("-")
            # 템플릿에 최대값 정보가 없으면 체크박스 비활성화
            for chk in self.max_checkboxes.values():
                chk.setEnabled(False)
        else:
            self.template_max_label.setText(str(max_val))
            for row, chk in self.max_checkboxes.items():
                chk.setEnabled(True)
                # 만약 이미 체크되어 있다면 스핀박스 값 업데이트 및 비활성화
                if chk.isChecked():
                    spin = self.product_table.cellWidget(row, 1)
                    if isinstance(spin, QSpinBox):
                        spin.setValue(max_val)
                        spin.setMaximum(max_val)
                        spin.setEnabled(False)
                else:
                    # 체크되지 않은 경우 스핀박스 최대값은 최소 보장 값 유지
                    spin = self.product_table.cellWidget(row, 1)
                    if isinstance(spin, QSpinBox):
                        # 원래 최대 999로 되돌리기
                        spin.setMaximum(999)
                        spin.setEnabled(True)

    def on_template_changed(self, index):
        # 템플릿 변경 시 라벨과 체크박스 상태 업데이트
        self.update_max_label_and_checkboxes()
        # 바코드 크기 자동 최적화
        self.template_auto_size(index)

    def template_auto_size(self, index):
        """템플릿 변경 시 셀 크기에 맞춰 바코드 크기를 자동 최적화 (MM 단위 통일)"""
        cell_size = None
        if self.cell_sizes is None:
            return

        # dict 형태일 때: 전체 경로 우선, basename 다음으로 시도
        if isinstance(self.cell_sizes, dict):
            selected_path = self.templates[index] if 0 <= index < len(self.templates) else None
            if selected_path and selected_path in self.cell_sizes:
                cell_size = self.cell_sizes[selected_path]
            else:
                base = os.path.basename(selected_path) if selected_path else None
                if base and base in self.cell_sizes:
                    cell_size = self.cell_sizes[base]

        # list/tuple 형태일 때: 인덱스 기반으로 찾기
        elif isinstance(self.cell_sizes, (list, tuple)):
            if 0 <= index < len(self.cell_sizes):
                cell_size = self.cell_sizes[index]
            else:
                # 경우에 따라 cell_sizes 자체가 (w,h) 한쌍일 수도 있음
                if len(self.cell_sizes) == 2 and all(isinstance(x, (int, float)) for x in self.cell_sizes):
                    cell_size = tuple(self.cell_sizes)

        # 성공적으로 셀 크기 정보를 얻었으면 바코드 크기를 최적화 (모든 값 MM 단위)
        try:
            if cell_size and len(cell_size) >= 2:
                cell_w_mm = float(cell_size[0])  # 셀 너비 (MM)
                cell_h_mm = float(cell_size[1])  # 셀 높이 (MM)

                # 바코드 최적 크기 계산 (MM 단위, 바코드 라이브러리 제약 고려)
                # 1. 모듈 너비: 셀 너비에서 여백 제외하고 바코드 길이로 나눔
                # 일반적인 Code128 바코드는 약 95개 모듈로 구성
                barcode_modules = 95
                margin_mm = 6.0  # 좌우 여백 총합 (quiet_zone 고려)
                available_width_mm = max(10.0, cell_w_mm - margin_mm)
                optimal_module_width_mm = available_width_mm / barcode_modules
                
                # 2. 모듈 높이: 셀 높이에서 텍스트 영역 제외
                text_area_mm = 6.0
                available_height_mm = max(8.0, cell_h_mm - text_area_mm)
                optimal_module_height_mm = available_height_mm
                
                # 3. 여백: 바코드 라이브러리 최소 요구사항 고려
                optimal_quiet_zone_mm = max(2.5, min(4.0, cell_w_mm * 0.08))
                
                # 바코드 라이브러리 제약에 맞는 안전한 범위로 제한
                final_module_width = max(0.2, min(1.0, optimal_module_width_mm))
                final_module_height = max(5.0, min(25.0, optimal_module_height_mm))
                final_quiet_zone = max(2.5, min(5.0, optimal_quiet_zone_mm))

                # 스핀박스에 반영 (소수 자릿수에 맞춰 반올림)
                self.module_width_spin.setValue(round(final_module_width, 2))
                self.module_height_spin.setValue(round(final_module_height, 1))
                self.quiet_zone_spin.setValue(round(final_quiet_zone, 1))
                
                print(f"바코드 크기 자동 최적화 완료:")
                print(f"  셀 크기: {cell_w_mm:.1f}mm x {cell_h_mm:.1f}mm")
                print(f"  모듈 너비: {final_module_width:.2f}mm")
                print(f"  모듈 높이: {final_module_height:.1f}mm")
                print(f"  좌우 여백: {final_quiet_zone:.1f}mm")
                
        except Exception as e:
            print(f"바코드 크기 자동 최적화 실패: {e}")
            # 실패 시 조용히 무시 (기존 값 유지)
    
    def _on_single_file_changed(self, state):
        """단일 파일 생성 체크박스 상태 변경 시 처리"""
        if state == Qt.CheckState.Checked:
            self.single_file_info.setVisible(True)
        else:
            self.single_file_info.setVisible(False)
    
    def on_max_checked(self, row, state):
        """특정 행의 Max 체크박스가 변경되었을 때 처리"""
        max_val = self.get_current_template_max()
        spin = self.product_table.cellWidget(row, 1)
        chk = self.max_checkboxes.get(row)
        if not isinstance(spin, QSpinBox) or chk is None:
            return

        if state == Qt.CheckState.Checked:
            if max_val is None:
                # 최대값 정보 없으면 바로 체크 해제
                chk.blockSignals(True)
                chk.setChecked(False)
                chk.blockSignals(False)
                return
            # 스핀박스를 최대값으로 설정하고 편집 비활성화
            spin.setMaximum(max_val)
            spin.setValue(max_val)
            spin.setEnabled(False)
        else:
            # 체크 해제: 스핀박스 편집 가능 및 기본 최대값 복원
            spin.setEnabled(True)
            spin.setMaximum(999)

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
            quantities[product.barcode_num] = spin_box.value()

        return {
            "template": selected_template_path,
            "quantities": quantities,
            "barcode_options": barcode_options,
            "max_label_count": self.template_max_label.text(),
            "single_file": self.single_file_checkbox.isChecked()
        }
