from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QListWidget,
    QPushButton,
    QMessageBox,
    QInputDialog,
    QListWidgetItem,
    QLabel,
    QSpinBox,
    QFormLayout,
    QComboBox,
    QDialogButtonBox,
)
from PyQt6.QtCore import pyqtSignal
from typing import List, Dict

from src.services.excel_service import ExcelService


class CategoryDialog(QDialog):
    """TYPE 관리 다이얼로그"""

    # 종류 목록이 변경되었음을 알리는 신호
    categories_updated = pyqtSignal()

    def __init__(self, excel_service: ExcelService, parent=None):
        super().__init__(parent)
        self.excel_service = excel_service

        self.setWindowTitle("TYPE 관리")
        self.setMinimumWidth(450)
        self.setMinimumHeight(400)

        self.setup_ui()
        self.load_type_ids()

    def setup_ui(self):
        """UI 구성"""
        layout = QVBoxLayout(self)

        # TYPE_ID 목록 (종류명: TYPE_ID 형식으로 표시)
        self.type_id_list = QListWidget()
        self.type_id_list.itemDoubleClicked.connect(self.edit_type_id)
        layout.addWidget(self.type_id_list)

        # TYPE_ID 관리 버튼들
        button_layout = QHBoxLayout()

        self.add_button = QPushButton("TYPE 추가")
        self.add_button.clicked.connect(self.add_type_id)

        self.edit_button = QPushButton("수정")
        self.edit_button.clicked.connect(self.edit_type_id)

        self.delete_button = QPushButton("삭제")
        self.delete_button.clicked.connect(self.delete_type_id)

        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.edit_button)
        button_layout.addWidget(self.delete_button)

        layout.addLayout(button_layout)

    def load_type_ids(self):
        """TYPE_ID 목록을 불러와 리스트에 표시"""
        self.type_id_list.clear()
        categories_dict = self.excel_service.get_categories()

        # 종류명: TYPE_ID 형식으로 표시 (TYPE_ID 순으로 정렬)
        for type_name, type_id in sorted(categories_dict.items(), key=lambda x: x[1]):
            display_text = f"{type_name}: {type_id}"
            item = QListWidgetItem(display_text)
            item.setData(1, {"type_name": type_name, "type_id": type_id})  # 데이터 저장
            self.type_id_list.addItem(item)

    def add_type_id(self):
        """새 TYPE 추가"""
        dialog = QDialog(self)
        dialog.setWindowTitle("새 TYPE 추가")
        layout = QFormLayout(dialog)

        # 종류 이름 입력
        type_name_input = QInputDialog()
        type_name, ok = QInputDialog.getText(
            self, "TYPE 이름", "TYPE 이름을 입력하세요:"
        )

        if not ok or not type_name:
            return

        type_name = type_name.strip()
        if type_name in self.excel_service.get_all_categories():
            QMessageBox.warning(self, "중복 오류", "이미 존재하는 TYPE 이름입니다.")
            return

        # TYPE_ID 입력
        existing_ids = list(self.excel_service.get_categories().values())
        max_id = max(existing_ids) if existing_ids else 0

        type_id, ok = QInputDialog.getInt(
            self,
            "TYPE_ID",
            f"TYPE_ID를 입력하세요 (기존 ID: {existing_ids}):",
            value=max_id + 1,
            min=0,
            max=9999,
        )

        if not ok:
            return

        # TYPE_ID 중복 확인
        if type_id in existing_ids:
            QMessageBox.warning(
                self, "중복 오류", f"TYPE_ID {type_id}는 이미 사용 중입니다."
            )
            return

        if self.excel_service.add_type_with_id(type_name, type_id):
            self.load_type_ids()
            self.categories_updated.emit()
            QMessageBox.information(
                self,
                "성공",
                f"'{type_name}' TYPE이 추가되었습니다.\nTYPE_ID: {type_id}",
            )
        else:
            QMessageBox.critical(self, "오류", "TYPE 추가에 실패했습니다.")

    def edit_type_id(self):
        """선택된 TYPE 수정 (이름 및 TYPE_ID)"""
        selected_item = self.type_id_list.currentItem()
        if not selected_item:
            QMessageBox.warning(self, "선택 오류", "수정할 항목을 선택하세요.")
            return

        item_data = selected_item.data(1)
        if not item_data:
            return

        old_type_name = item_data["type_name"]
        current_type_id = item_data["type_id"]

        # 수정 다이얼로그 생성
        dialog = QDialog(self)
        dialog.setWindowTitle("TYPE 수정")
        layout = QFormLayout(dialog)

        # TYPE 이름 입력 필드
        from PyQt6.QtWidgets import QLineEdit

        name_input = QLineEdit(old_type_name)
        layout.addRow("TYPE 이름:", name_input)

        # TYPE_ID 입력 필드
        id_input = QSpinBox()
        id_input.setRange(0, 9999)
        id_input.setValue(current_type_id)
        layout.addRow("TYPE_ID:", id_input)

        # 버튼
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_type_name = name_input.text().strip()
            new_type_id = id_input.value()

            # 변경사항 확인
            name_changed = new_type_name != old_type_name
            id_changed = new_type_id != current_type_id

            if not name_changed and not id_changed:
                return

            # 이름 중복 확인
            if (
                name_changed
                and new_type_name in self.excel_service.get_all_categories()
            ):
                QMessageBox.warning(self, "중복 오류", "이미 존재하는 TYPE 이름입니다.")
                return

            # TYPE_ID 중복 확인
            if id_changed:
                existing_categories = self.excel_service.get_categories()
                if new_type_id in existing_categories.values():
                    QMessageBox.warning(
                        self,
                        "중복 오류",
                        f"TYPE_ID {new_type_id}는 이미 사용 중입니다.",
                    )
                    return

            # 수정 실행
            success = True
            if name_changed:
                success = success and self.excel_service.update_type_name(
                    old_type_name, new_type_name
                )

            if id_changed and success:
                # 이름이 변경된 경우 새 이름 사용, 아니면 기존 이름 사용
                target_name = new_type_name if name_changed else old_type_name
                success = success and self.excel_service.update_type_id(
                    target_name, new_type_id
                )

            if success:
                self.load_type_ids()
                self.categories_updated.emit()
                QMessageBox.information(
                    self,
                    "성공",
                    f"TYPE이 수정되었습니다.\n이름: {old_type_name} → {new_type_name}\nTYPE_ID: {current_type_id} → {new_type_id}",
                )
            else:
                QMessageBox.critical(self, "오류", "TYPE 수정에 실패했습니다.")

    def delete_type_id(self):
        """선택된 TYPE 삭제"""
        selected_item = self.type_id_list.currentItem()
        if not selected_item:
            QMessageBox.warning(self, "선택 오류", "삭제할 항목을 선택하세요.")
            return

        item_data = selected_item.data(1)
        if not item_data:
            return

        type_name = item_data["type_name"]
        type_id = item_data["type_id"]

        # 해당 TYPE이 사용 중인지 확인
        if self.excel_service.is_type_name_in_use(type_name):
            QMessageBox.warning(
                self,
                "삭제 불가",
                f"'{type_name}' TYPE은 현재 하나 이상의 상품에서 사용 중이므로 삭제할 수 없습니다.",
            )
            return

        reply = QMessageBox.question(
            self,
            "TYPE 삭제 확인",
            f"'{type_name}' (TYPE_ID: {type_id})를 정말 삭제하시겠습니까?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )

        if reply == QMessageBox.StandardButton.Yes:
            if self.excel_service.delete_type_name(type_name):
                self.load_type_ids()
                self.categories_updated.emit()
                QMessageBox.information(
                    self, "성공", f"'{type_name}' TYPE이 삭제되었습니다."
                )
            else:
                QMessageBox.critical(self, "오류", "TYPE 삭제에 실패했습니다.")
