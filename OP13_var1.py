from __future__ import annotations

import sys
from dataclasses import dataclass
from typing import Optional

from PySide6.QtCore import Qt, QDate, QLocale, QPersistentModelIndex, QStringListModel, Signal
from PySide6.QtGui import QIntValidator
from PySide6.QtWidgets import (
    QApplication,
    QCompleter,
    QComboBox,
    QDateEdit,
    QDoubleSpinBox,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QStyledItemDelegate,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


@dataclass
class SpiceRow:
    name: str = ""
    code: str = ""
    opening_balance: float = 0.0
    received: float = 0.0
    closing_balance: float = 0.0
    consumed: float = 0.0


POSITIONS_APPROVE = [
    "Генеральный директор",
    "Директор по производству",
    "Главный бухгалтер",
]
POSITIONS_COMPILER = [
    "Кладовщик",
    "Заведующий складом",
]

ORGANIZATIONS = {
    "ООО Ресторан 'Пряный Вкус'": "10293847",
    "АО Столовая 'Восток'": "56473829",
    "ООО Кафе 'Специи и Соль'": "12345678",
    "АО Фабрика-кухня 'Гурман'": "87654321",
    "ЗАО Мясокомбинат 'СпецПром'": "11223344",
    "ПАО Общепит 'Соляной Рай'": "55667788",
    "ООО Пекарня 'Аромат'": "99887766",
    "АО Кондитерская Фабрика 'Ваниль'": "44556677",
    "ООО Кейтеринг 'Пряности'": "22334455",
    "ПАО Столовая Завода 'СольТрейд'": "66778899",
    "АО Ресторанный Двор 'Перец'": "33445566",
    "ООО Кулинарный Цех 'СпецииПро'": "77889900",
    "ЗАО Рыбокомбинат 'Морская Соль'": "99001122",
    "ПАО Кафе 'Чесночный Аромат'": "11220099",
    "ООО Общественное Питание 'Соль и Перец'": "55443322",
}

DEPARTMENTS = {
    "Склад специй и приправ": "51.17.10",
    "Склад соли и консервантов": "51.17.20",
    "Горячий цех": "15.11.10",
    "Холодный цех": "15.11.20",
    "Кондитерский цех": "15.82.10",
    "Мясной цех": "15.11.30",
    "Рыбный цех": "15.20.10",
    "Цех полуфабрикатов": "15.13.10",
    "Производственный цех (общий)": "15.11.00",
    "Склад сырья и ингредиентов": "51.17.00",
    "Отдел контроля качества и расхода": "74.30.10",
    "Кухня ресторана": "55.30.10",
    "Кухня столовой": "55.51.10",
    "Кейтеринг-цех": "55.52.10",
    "Пекарный цех": "15.81.10",
    "Отдел снабжения специями": "51.38.10",
    "Склад готовой продукции": "52.10.10",
}

SPICES = {
    "Перец черный молотый": "091011",
    "Паприка": "090422",
    "Кориандр молотый": "090921",
    "Соль поваренная": "1541010",
    "Лавровый лист": "091099",
    "Зира": "090930",
    "Карри порошок": "091091",
    "Куркума молотая": "091030",
    "Имбирь молотый": "091010",
    "Чеснок сушеный": "071290",
    "Укроп сушеный": "091099",
    "Петрушка сушеная": "091099",
    "Базилик сушеный": "091099",
    "Розмарин сушеный": "091099",
    "Тимьян": "091099",
    "Гвоздика": "090700",
    "Корица молотая": "090610",
    "Мускатный орех молотый": "090810",
    "Ваниль": "090500",
    "Перец красный молотый": "090420",
    "Соль морская": "1541020",
    "Соль йодированная": "1541030",
    "Горчица порошок": "090920",
    "Анис": "090961",
    "Фенхель": "090962",
    "Перец душистый": "090411",
    "Хмели-сунели": "091091",
}

POSITION_TO_PERSONS: dict[str, list[str]] = {
    "Генеральный директор": ["Иванов Иван Иванович"],
    "Директор по производству": ["Смирнов Алексей Петрович"],
    "Главный бухгалтер": ["Кузнецова Мария Сергеевна"],
    "Кладовщик": ["Соколов Дмитрий Андреевич", "Николаев Игорь Сергеевич"],
    "Заведующий складом": ["Васильева Ольга Николаевна", "Федорова Елена Ивановна"],
    "Бухгалтер": ["Петрова Анна Сергеевна", "Егорова Наталья Викторовна"],
}

PERSON_TO_POSITION = {
    person: position
    for position, people in POSITION_TO_PERSONS.items()
    for person in people
}

NAMES_APPROVE = [POSITION_TO_PERSONS[pos][0] for pos in POSITIONS_APPROVE if POSITION_TO_PERSONS.get(pos)]
NAMES_COMPILER = [person for pos in POSITIONS_COMPILER for person in POSITION_TO_PERSONS.get(pos, [])]
NAMES_ACCOUNTANT = POSITION_TO_PERSONS["Бухгалтер"]


def full_name_to_signature(full_name: str) -> str:
    parts = full_name.strip().split()
    if not parts:
        return ""
    initials = "".join(f"{part[0]}." for part in parts[1:] if part)
    return f"{parts[0]} {initials}".strip()


class MoneySpin(QDoubleSpinBox):
    valueChangedSafe = Signal()

    def __init__(self, parent: Optional[QWidget] = None, maximum: float = 1e9):
        super().__init__(parent)
        self.setDecimals(2)
        self.setRange(0.0, maximum)
        self.setSingleStep(1.00)
        self.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.valueChanged.connect(lambda *_: self.valueChangedSafe.emit())


class SpiceCompleterDelegate(QStyledItemDelegate):
    def __init__(self, spice_map: dict[str, str], parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._spice_to_code = spice_map
        self._code_to_spice = {code: spice for spice, code in spice_map.items()}
        self._spice_lower = {spice.casefold(): spice for spice in spice_map}
        self._code_lower = {code.casefold(): code for code in self._code_to_spice}

    def createEditor(self, parent: QWidget, option, index):
        if index.column() not in (1, 2):
            return super().createEditor(parent, option, index)

        persistent_index = QPersistentModelIndex(index)
        if index.column() == 1:
            editor = QComboBox(parent)
            editor.setEditable(True)
            editor.addItems(list(self._spice_to_code.keys()))
            editor.setCurrentText(str(index.data(Qt.ItemDataRole.EditRole) or ""))

            line_edit = editor.lineEdit()
            if line_edit is not None:
                model = QStringListModel(list(self._spice_to_code.keys()), line_edit)
                completer = QCompleter(model, line_edit)
                completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
                completer.setFilterMode(Qt.MatchFlag.MatchContains)
                completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
                line_edit.setCompleter(completer)
                completer.activated.connect(lambda text, idx=persistent_index: self._apply_pair_value(idx, text))
                line_edit.returnPressed.connect(
                    lambda idx=persistent_index, cb=editor: self._apply_pair_value(idx, cb.currentText())
                )

            editor.textActivated.connect(lambda text, idx=persistent_index: self._apply_pair_value(idx, text))
            return editor

        editor = QLineEdit(parent)
        model = QStringListModel(list(self._code_to_spice.keys()), editor)
        completer = QCompleter(model, editor)
        completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        completer.setFilterMode(Qt.MatchFlag.MatchContains)
        completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
        editor.setCompleter(completer)
        completer.activated.connect(lambda text, idx=persistent_index: self._apply_pair_value(idx, text))
        editor.returnPressed.connect(lambda idx=persistent_index, e=editor: self._apply_pair_value(idx, e.text()))
        return editor

    def setModelData(self, editor, model, index) -> None:
        if isinstance(editor, QComboBox) and index.column() == 1:
            raw_value = editor.currentText().strip()
            spice = self._spice_lower.get(raw_value.casefold())
            if spice:
                model.setData(index, spice, Qt.ItemDataRole.EditRole)
                model.setData(model.index(index.row(), 2), self._spice_to_code[spice], Qt.ItemDataRole.EditRole)
            else:
                model.setData(index, raw_value, Qt.ItemDataRole.EditRole)
            return
        super().setModelData(editor, model, index)

    def _apply_pair_value(self, index: QPersistentModelIndex, text: str) -> None:
        if not index.isValid():
            return
        model = index.model()

        if index.column() == 1:
            spice = self._spice_lower.get(text.casefold())
            if not spice:
                return
            code = self._spice_to_code[spice]
            model.setData(index, spice, Qt.ItemDataRole.EditRole)
            model.setData(model.index(index.row(), 2), code, Qt.ItemDataRole.EditRole)
            return

        code = self._code_lower.get(text.casefold())
        if not code:
            return
        spice = self._code_to_spice[code]
        model.setData(index, code, Qt.ItemDataRole.EditRole)
        model.setData(model.index(index.row(), 1), spice, Qt.ItemDataRole.EditRole)


class OP13BlankWindow(QMainWindow):
    DATA_ROWS_DEFAULT = 3

    def __init__(self):
        super().__init__()
        self.setWindowTitle("ОП-13 - Акт расхода специй и соли (вариант 1)")
        self.resize(850, 950)
        self.setMinimumSize(850, 950)

        self._did_first_table_resize = False

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop)
        #scroll.setStyleSheet("QScrollArea { background-color: white; border: none; }")
        #scroll.viewport().setStyleSheet("background-color: white;")
        root = QWidget()
        root.setMaximumWidth(820)
        scroll.setWidget(root)
        self.setCentralWidget(scroll)

        self._root_layout = QVBoxLayout(root)
        self._root_layout.setContentsMargins(14, 14, 14, 14)
        self._root_layout.setSpacing(14)

        self._root_layout.addWidget(self._build_title_block())
        self._root_layout.addWidget(self._build_requisites_block())
        self._root_layout.addWidget(self._build_table_block())
        self._root_layout.addWidget(self._build_reference_block())
        self._root_layout.addWidget(self._build_signatures_block())
        self._root_layout.addWidget(self._build_bottom_buttons_block())

        self._root_layout.addStretch(1)

        self._apply_ru_locale()
        self._setup_requisites_completers()
        self._set_default_reporting_period()
        self._syncing_person_fields = False
        self._initialize_people_selectors()
        self._recalc_reference()

    def showEvent(self, event):
        super().showEvent(event)
        if not self._did_first_table_resize:
            self._did_first_table_resize = True
            self._adjust_table_height_to_rows()

    def _make_date_edit(self, display_format: str) -> QDateEdit:
        ed = QDateEdit()
        ed.setCalendarPopup(True)
        ed.setDisplayFormat(display_format)
        ed.setDate(QDate.currentDate())
        return ed

    def _set_default_reporting_period(self) -> None:
        first_day_prev_month = QDate.currentDate().addMonths(-1)
        first_day_prev_month = QDate(first_day_prev_month.year(), first_day_prev_month.month(), 1)
        last_day_prev_month = first_day_prev_month.addMonths(1).addDays(-1)
        self.ed_period_from.setDate(first_day_prev_month)
        self.ed_period_to.setDate(last_day_prev_month)

    def _make_completer(self, values: list[str], parent: QWidget) -> QCompleter:
        model = QStringListModel(values, parent)
        completer = QCompleter(model, parent)
        completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        completer.setFilterMode(Qt.MatchFlag.MatchContains)
        completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
        return completer

    def _find_match(self, text: str, values: list[str]) -> str:
        normalized = text.strip().casefold()
        for value in values:
            if value.casefold() == normalized:
                return value
        return ""

    def _setup_bidirectional_completion(
        self,
        left_edit: QLineEdit,
        right_edit: QLineEdit,
        left_to_right: dict[str, str],
    ) -> None:
        right_to_left = {v: k for k, v in left_to_right.items()}
        left_values = list(left_to_right.keys())
        right_values = list(right_to_left.keys())

        left_completer = self._make_completer(left_values, left_edit)
        right_completer = self._make_completer(right_values, right_edit)
        left_edit.setCompleter(left_completer)
        right_edit.setCompleter(right_completer)

        def apply_left(text: str) -> None:
            left_match = self._find_match(text, left_values)
            if not left_match:
                return
            right_value = left_to_right[left_match]
            left_edit.blockSignals(True)
            right_edit.blockSignals(True)
            left_edit.setText(left_match)
            right_edit.setText(right_value)
            right_edit.blockSignals(False)
            left_edit.blockSignals(False)

        def apply_right(text: str) -> None:
            right_match = self._find_match(text, right_values)
            if not right_match:
                return
            left_value = right_to_left[right_match]
            right_edit.blockSignals(True)
            left_edit.blockSignals(True)
            right_edit.setText(right_match)
            left_edit.setText(left_value)
            left_edit.blockSignals(False)
            right_edit.blockSignals(False)

        left_completer.activated.connect(apply_left)
        right_completer.activated.connect(apply_right)
        left_edit.returnPressed.connect(lambda: apply_left(left_edit.text()))
        right_edit.returnPressed.connect(lambda: apply_right(right_edit.text()))

    def _setup_requisites_completers(self) -> None:
        self._setup_bidirectional_completion(self.ed_org, self.ed_okpo, ORGANIZATIONS)
        self._setup_bidirectional_completion(self.ed_dept, self.ed_okdp, DEPARTMENTS)

    def _replace_combo_items(self, combo: QComboBox, items: list[str], selected_text: str = "") -> None:
        combo.blockSignals(True)
        combo.clear()
        combo.addItem("")
        combo.addItems(items)

        if selected_text and selected_text in items:
            combo.setCurrentText(selected_text)
        elif items:
            combo.setCurrentIndex(1)
        else:
            combo.setCurrentIndex(0)
        combo.blockSignals(False)

    def _initialize_people_selectors(self) -> None:
        self._replace_combo_items(self.cb_head_name, NAMES_APPROVE)
        self._replace_combo_items(self.cb_compiler_name, NAMES_COMPILER)
        self._replace_combo_items(self.cb_accountant_name, NAMES_ACCOUNTANT)
        self.cb_head_name.setCurrentIndex(0)
        self.cb_compiler_name.setCurrentIndex(0)
        # Бухгалтер всегда выбран — проставляем первого человека из списка
        if NAMES_ACCOUNTANT:
            self.cb_accountant_name.setCurrentText(NAMES_ACCOUNTANT[0])
            self.ed_accountant_signature.setText(
                full_name_to_signature(NAMES_ACCOUNTANT[0])
            )

    def _build_title_block(self) -> QWidget:
        box = QGroupBox("Унифицированная форма № ОП-13 - Акт расхода специй и соли")
        layout = QVBoxLayout(box)
        layout.setSpacing(10)

        row = QHBoxLayout()
        row.setSpacing(10)

        self.ed_doc_no = QLineEdit()
        self.ed_doc_no.setObjectName("doc_no")
        self.ed_doc_no.setMaximumWidth(100)
        self.ed_doc_no.setPlaceholderText("Номер")

        self.ed_doc_date = self._make_date_edit("dd.MM.yyyy")
        self.ed_doc_date.setObjectName("doc_date")
        self.ed_doc_date.setMaximumWidth(120)

        self.ed_period_from = self._make_date_edit("dd.MM.yyyy")
        self.ed_period_from.setObjectName("period_from")
        self.ed_period_from.setMaximumWidth(120)

        self.ed_period_to = self._make_date_edit("dd.MM.yyyy")
        self.ed_period_to.setObjectName("period_to")
        self.ed_period_to.setMaximumWidth(120)

        row.addWidget(QLabel("Номер документа"))
        row.addWidget(self.ed_doc_no)
        row.addWidget(QLabel("Дата составления"))
        row.addWidget(self.ed_doc_date)
        row.addWidget(QLabel("Отчетный период с"))
        row.addWidget(self.ed_period_from)
        row.addWidget(QLabel("по"))
        row.addWidget(self.ed_period_to)
        row.addStretch(1)
        layout.addLayout(row)

        return box

    def _build_requisites_block(self) -> QWidget:
        box = QGroupBox("Реквизиты")
        root = QVBoxLayout(box)
        root.setSpacing(6)

        self.ed_org = QLineEdit()
        self.ed_org.setObjectName("organization")
        self.ed_org.setPlaceholderText("Наименование организации")
        self.ed_org.setMinimumWidth(220)

        self.ed_dept = QLineEdit()
        self.ed_dept.setObjectName("department")
        self.ed_dept.setPlaceholderText("Структурное подразделение")
        self.ed_dept.setMinimumWidth(220)

        self.ed_okpo = QLineEdit()
        self.ed_okpo.setObjectName("okpo")
        self.ed_okpo.setValidator(QIntValidator(0, 999999999))
        self.ed_okpo.setPlaceholderText("ОКПО")
        self.ed_okpo.setFixedWidth(80)

        self.ed_okdp = QLineEdit()
        self.ed_okdp.setObjectName("okdp")
        self.ed_okdp.setPlaceholderText("ОКДП")
        self.ed_okdp.setFixedWidth(80)

        self.ed_operation = QLineEdit()
        self.ed_operation.setObjectName("operation_kind")
        self.ed_operation.setPlaceholderText("Вид")

        approve_box = QGroupBox("УТВЕРЖДАЮ")
        approve_box.setMaximumWidth(350)
        approve_grid = QGridLayout(approve_box)
        approve_grid.setHorizontalSpacing(10)
        approve_grid.setVerticalSpacing(6)

        self.cb_head_position = QComboBox()
        self.cb_head_position.setObjectName("cb_head_position")
        self.cb_head_position.setEditable(True)
        self.cb_head_position.addItem("")
        self.cb_head_position.addItems(POSITIONS_APPROVE)

        self.cb_head_name = QComboBox()
        self.cb_head_name.setObjectName("cb_head_name")
        self.cb_head_name.setEditable(True)
        self.cb_head_name.addItem("")
        self.cb_head_name.addItems(NAMES_APPROVE)
        self.cb_head_position.currentTextChanged.connect(self._on_head_position_changed)
        self.cb_head_name.currentTextChanged.connect(self._on_head_name_changed)

        self.ed_act_date = self._make_date_edit("«dd» MMMM yyyy 'г.'")
        self.ed_act_date.setObjectName("approval_date")

        approve_grid.addWidget(QLabel("Должность:"), 0, 0)
        approve_grid.addWidget(self.cb_head_position, 0, 1)
        approve_grid.addWidget(QLabel("Расшифровка:"), 1, 0)
        approve_grid.addWidget(self.cb_head_name, 1, 1)
        approve_grid.addWidget(QLabel("Дата утверждения:"), 2, 0)
        approve_grid.addWidget(self.ed_act_date, 2, 1)
        approve_grid.setColumnStretch(1, 1)

        top_row = QHBoxLayout()
        top_row.setSpacing(8)

        left_grid = QGridLayout()
        left_grid.setHorizontalSpacing(4)
        left_grid.setVerticalSpacing(6)
        left_grid.addWidget(QLabel("Организация:"), 0, 0)
        left_grid.addWidget(self.ed_org, 0, 1, 1, 5)
        left_grid.addWidget(QLabel("Подразделение:"), 1, 0)
        left_grid.addWidget(self.ed_dept, 1, 1, 1, 5)
        left_grid.addWidget(QLabel("ОКДП:"), 2, 0)
        left_grid.addWidget(self.ed_okdp, 2, 1)
        left_grid.addWidget(QLabel("ОКПО:"), 2, 2)
        left_grid.addWidget(self.ed_okpo, 2, 3)
        left_grid.addWidget(QLabel("Вид операции:"), 2, 4)
        left_grid.addWidget(self.ed_operation, 2, 5)
        left_grid.setColumnMinimumWidth(0, 100)
        left_grid.setColumnStretch(1, 1)
        left_grid.setColumnStretch(3, 0)
        left_grid.setColumnStretch(5, 1)

        top_row.addLayout(left_grid, 1)
        top_row.addWidget(approve_box, 0)

        root.addLayout(top_row)

        return box

    def _adjust_table_height_to_rows(self) -> None:
        hh = self.table.horizontalHeader()
        header_h = max(hh.height(), hh.sizeHint().height())

        h = header_h
        for r in range(self.table.rowCount()):
            h += self.table.rowHeight(r)

        h += 2 * self.table.frameWidth()
        self.table.setFixedHeight(h)

    def _setup_totals_row(self, preserve_values: bool = True) -> None:
        self.table.clearSpans()
        totals_row = self.table.rowCount() - 1

        prev = (0.0, 0.0, 0.0, 0.0)
        if preserve_values and hasattr(self, "ed_total_open"):
            try:
                prev = (
                    self.ed_total_open.value(),
                    self.ed_total_recv.value(),
                    self.ed_total_close.value(),
                    self.ed_total_cons.value(),
                )
            except RuntimeError:
                prev = (0.0, 0.0, 0.0, 0.0)

        itogo_item = self.table.item(totals_row, 0)
        if itogo_item is None:
            itogo_item = QTableWidgetItem()
            self.table.setItem(totals_row, 0, itogo_item)
        itogo_item.setText("Итого")
        itogo_item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
        itogo_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)

        self.table.setSpan(totals_row, 0, 1, 3)

        def mk_total_spin() -> MoneySpin:
            w = MoneySpin()
            w.setButtonSymbols(QDoubleSpinBox.ButtonSymbols.NoButtons)
            w.setFrame(False)
            w.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            return w

        self.ed_total_open = mk_total_spin()
        self.ed_total_open.setObjectName("total_opening")

        self.ed_total_recv = mk_total_spin()
        self.ed_total_recv.setObjectName("total_received")

        self.ed_total_close = mk_total_spin()
        self.ed_total_close.setObjectName("total_closing")

        self.ed_total_cons = mk_total_spin()
        self.ed_total_cons.setObjectName("total_consumed")

        self.ed_total_open.setValue(prev[0])
        self.ed_total_recv.setValue(prev[1])
        self.ed_total_close.setValue(prev[2])
        self.ed_total_cons.setValue(prev[3])

        self.table.setCellWidget(totals_row, 3, self.ed_total_open)
        self.table.setCellWidget(totals_row, 4, self.ed_total_recv)
        self.table.setCellWidget(totals_row, 5, self.ed_total_close)
        self.table.setCellWidget(totals_row, 6, self.ed_total_cons)

        self.table.setRowHeight(totals_row, 32)

    def _build_table_block(self) -> QWidget:
        box = QGroupBox("Расход специй и соли")
        layout = QVBoxLayout(box)
        layout.setSpacing(10)

        self.table = QTableWidget(self.DATA_ROWS_DEFAULT + 1, 7)
        self.table.setObjectName("spices_table")
        self.table.setHorizontalHeaderLabels([
            "№",
            "Наименование",
            "Код",
            "Остаток на начало\n(сумма), руб.коп.",
            "Поступило за период\n(сумма), руб.коп.",
            "Остаток на конец\n(сумма), руб.коп.",
            "Израсходовано\n(сумма), руб.коп.",
        ])
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)

        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        self._renumber_table()
        self._setup_totals_row(preserve_values=False)
        self._spice_delegate = SpiceCompleterDelegate(SPICES, self.table)
        self.table.setItemDelegateForColumn(1, self._spice_delegate)
        self.table.setItemDelegateForColumn(2, self._spice_delegate)

        hdr = self.table.horizontalHeader()
        hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)

        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(0, 40)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(1, 160)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(2, 65)
        for c in range(3, 7):
            hdr.setSectionResizeMode(c, QHeaderView.ResizeMode.Stretch)

        layout.addWidget(self.table)
        self._adjust_table_height_to_rows()

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        self.btn_add_row = QPushButton("Добавить строку")
        self.btn_del_row = QPushButton("Удалить выбранную")
        self.btn_add_row.clicked.connect(self._add_table_row)
        self.btn_del_row.clicked.connect(self._delete_selected_rows)
        btn_row.addWidget(self.btn_add_row)
        btn_row.addWidget(self.btn_del_row)
        layout.addLayout(btn_row)

        return box

    def _build_reference_block(self) -> QWidget:
        box = QGroupBox("Справка о стоимости специй и соли, включенной в калькуляцию блюд")
        grid = QGridLayout(box)
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(8)

        grid.setColumnStretch(0, 3)
        grid.setColumnStretch(1, 3)
        grid.setColumnStretch(2, 3)
        grid.setColumnStretch(3, 3)

        def mk_label(text: str, bold: bool = False) -> QLabel:
            lbl = QLabel(f"<b>{text}</b>" if bold else text)
            lbl.setWordWrap(True)
            return lbl

        def expand_fixed(w: QWidget) -> QWidget:
            w.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            return w

        self.sb_spice_dishes = QSpinBox()
        self.sb_spice_dishes.setObjectName("spice_dishes_qty")
        self.sb_spice_dishes.setRange(0, 10_000_000)

        self.sp_spice_per_dish = expand_fixed(MoneySpin(maximum=1_000_000))
        self.sp_spice_per_dish.setObjectName("spice_price_per_dish")

        self.sp_spice_sum = expand_fixed(MoneySpin(maximum=1_000_000_000))
        self.sp_spice_sum.setObjectName("spice_sum")
        self.sp_spice_sum.setReadOnly(True)
        self.sp_spice_sum.setButtonSymbols(QDoubleSpinBox.ButtonSymbols.NoButtons)

        self.sb_salt_dishes = QSpinBox()
        self.sb_salt_dishes.setObjectName("salt_dishes_qty")
        self.sb_salt_dishes.setRange(0, 10_000_000)

        self.sp_salt_per_dish = expand_fixed(MoneySpin(maximum=1_000_000))
        self.sp_salt_per_dish.setObjectName("salt_price_per_dish")

        self.sp_salt_sum = expand_fixed(MoneySpin(maximum=1_000_000_000))
        self.sp_salt_sum.setObjectName("salt_sum")
        self.sp_salt_sum.setReadOnly(True)
        self.sp_salt_sum.setButtonSymbols(QDoubleSpinBox.ButtonSymbols.NoButtons)

        self.sp_ref_total = expand_fixed(MoneySpin(maximum=1_000_000_000))
        self.sp_ref_total.setObjectName("ref_total")
        self.sp_ref_total.setReadOnly(True)
        self.sp_ref_total.setButtonSymbols(QDoubleSpinBox.ButtonSymbols.NoButtons)

        self.sp_control_consumed = expand_fixed(MoneySpin(maximum=1_000_000_000))
        self.sp_control_consumed.setObjectName("control_consumed")

        self.sp_underspend = expand_fixed(MoneySpin(maximum=1_000_000_000))
        self.sp_underspend.setObjectName("underspend")
        self.sp_underspend.setReadOnly(True)
        self.sp_underspend.setButtonSymbols(QDoubleSpinBox.ButtonSymbols.NoButtons)

        grid.addWidget(mk_label("Показатель", bold=True), 0, 0)
        grid.addWidget(mk_label("Количество блюд", bold=True), 0, 1)
        grid.addWidget(mk_label("руб./коп. на блюдо", bold=True), 0, 2)
        grid.addWidget(mk_label("Сумма, руб./коп.", bold=True), 0, 3)

        grid.addWidget(mk_label("Продано блюд, в которые включена стоимость специй"), 1, 0)
        grid.addWidget(self.sb_spice_dishes, 1, 1)
        grid.addWidget(self.sp_spice_per_dish, 1, 2)
        grid.addWidget(self.sp_spice_sum, 1, 3)

        grid.addWidget(mk_label("Продано блюд, в которые включена стоимость соли"), 2, 0)
        grid.addWidget(self.sb_salt_dishes, 2, 1)
        grid.addWidget(self.sp_salt_per_dish, 2, 2)
        grid.addWidget(self.sp_salt_sum, 2, 3)

        grid.addWidget(mk_label("Итого", bold=True), 3, 0, 1, 3, alignment=Qt.AlignmentFlag.AlignRight)
        grid.addWidget(self.sp_ref_total, 3, 3)

        grid.addWidget(
            mk_label("Израсходовано согласно<br>контрольного расчета:"),
            4,
            0,
            1,
            3,
            alignment=Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
        )
        grid.addWidget(self.sp_control_consumed, 4, 3)

        grid.addWidget(mk_label("Сумма недорасхода:"), 5, 0, 1, 3, alignment=Qt.AlignmentFlag.AlignRight)
        grid.addWidget(self.sp_underspend, 5, 3)

        self.sb_spice_dishes.valueChanged.connect(self._recalc_reference)
        self.sb_salt_dishes.valueChanged.connect(self._recalc_reference)
        self.sp_spice_per_dish.valueChangedSafe.connect(self._recalc_reference)
        self.sp_salt_per_dish.valueChangedSafe.connect(self._recalc_reference)
        self.sp_control_consumed.valueChangedSafe.connect(self._recalc_reference)

        return box

    def _build_signatures_block(self) -> QWidget:
        box = QGroupBox("Подписи")
        grid = QGridLayout(box)
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(8)

        self.cb_compiler_position = QComboBox()
        self.cb_compiler_position.setObjectName("cb_compiler_position")
        self.cb_compiler_position.setEditable(True)
        self.cb_compiler_position.addItem("")
        self.cb_compiler_position.addItems(POSITIONS_COMPILER)

        self.ed_compiler_signature = QLineEdit()
        self.ed_compiler_signature.setObjectName("compiler_signature")
        self.ed_compiler_signature.setPlaceholderText("Подпись")
        self.ed_compiler_signature.setReadOnly(True)
        self.ed_compiler_signature.setMinimumWidth(140)

        self.cb_compiler_name = QComboBox()
        self.cb_compiler_name.setObjectName("cb_compiler_name")
        self.cb_compiler_name.setEditable(True)
        self.cb_compiler_name.addItem("")
        self.cb_compiler_name.addItems(NAMES_COMPILER)
        self.cb_compiler_position.currentTextChanged.connect(self._on_compiler_position_changed)
        self.cb_compiler_name.currentTextChanged.connect(self._on_compiler_name_changed)

        self.ed_accountant_signature = QLineEdit()
        self.ed_accountant_signature.setObjectName("accountant_signature")
        self.ed_accountant_signature.setPlaceholderText("Подпись")
        self.ed_accountant_signature.setReadOnly(True)
        self.ed_accountant_signature.setMinimumWidth(140)

        self.cb_accountant_name = QComboBox()
        self.cb_accountant_name.setObjectName("cb_accountant_name")
        self.cb_accountant_name.setEditable(True)
        self.cb_accountant_name.addItem("")
        self.cb_accountant_name.addItems(NAMES_ACCOUNTANT)
        self.cb_accountant_name.currentTextChanged.connect(self._on_accountant_name_changed)

        self.ed_accountant_position = QLineEdit("Бухгалтер")
        self.ed_accountant_position.setObjectName("accountant_position")
        self.ed_accountant_position.setReadOnly(True)

        grid.addWidget(QLabel("Расчет и справку составил:"), 0, 0)
        grid.addWidget(self.cb_compiler_position, 0, 1)
        grid.addWidget(self.ed_compiler_signature, 0, 2)
        grid.addWidget(self.cb_compiler_name, 0, 3)

        grid.addWidget(QLabel(""), 1, 0)
        grid.addWidget(self.ed_accountant_position, 1, 1)
        grid.addWidget(self.ed_accountant_signature, 1, 2)
        grid.addWidget(self.cb_accountant_name, 1, 3)

        grid.setColumnStretch(1, 2)
        grid.setColumnStretch(3, 2)

        return box

    def _build_bottom_buttons_block(self) -> QWidget:
        box = QWidget()
        row = QHBoxLayout(box)
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(10)

        self.btn_clear_form = QPushButton("Очистить форму")
        self.btn_clear_form.setObjectName("btn_clear_form")
        self.btn_clear_form.setStyleSheet(
            "QPushButton { background-color: #dc2626; color: white; border-radius: 4px; padding: 6px 16px; }"
            "QPushButton:hover { background-color: #b91c1c; }"
            "QPushButton:pressed { background-color: #991b1b; }"
        )
        self.btn_clear_form.clicked.connect(self.clear_form)

        self.btn_print = QPushButton("Печать")
        self.btn_print.setObjectName("btn_print")
        self.btn_print.setStyleSheet(
            "QPushButton { background-color: #2563eb; color: white; border-radius: 4px; padding: 6px 16px; }"
            "QPushButton:hover { background-color: #1d4ed8; }"
            "QPushButton:pressed { background-color: #1e40af; }"
        )

        self.btn_export_xls = QPushButton("Экспорт в XLS")
        self.btn_export_xls.setObjectName("btn_export_xls")
        self.btn_export_xls.setStyleSheet(
            "QPushButton { background-color: #16a34a; color: white; border-radius: 4px; padding: 6px 16px; }"
            "QPushButton:hover { background-color: #15803d; }"
            "QPushButton:pressed { background-color: #166534; }"
        )

        row.addWidget(self.btn_clear_form)
        row.addStretch(1)
        row.addWidget(self.btn_print)
        row.addWidget(self.btn_export_xls)

        return box

    def _apply_ru_locale(self) -> None:
        ru = QLocale(QLocale.Language.Russian, QLocale.Country.Russia)
        QLocale.setDefault(ru)
        for w in (self.ed_doc_date, self.ed_period_from, self.ed_period_to, self.ed_act_date):
            w.setLocale(ru)

    def _renumber_table(self) -> None:
        data_rows = self.table.rowCount() - 1
        for r in range(data_rows):
            item = self.table.item(r, 0)
            if item is None:
                item = QTableWidgetItem()
                self.table.setItem(r, 0, item)
            item.setText(str(r + 1))
            item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)

            for c in (1, 2):
                cell_item = self.table.item(r, c)
                if cell_item is None:
                    self.table.setItem(r, c, QTableWidgetItem(""))

    def _add_table_row(self) -> None:
        insert_at = self.table.rowCount() - 1
        self.table.insertRow(insert_at)
        self._renumber_table()
        self._setup_totals_row(preserve_values=True)
        self._adjust_table_height_to_rows()

    def _delete_selected_rows(self) -> None:
        totals_row = self.table.rowCount() - 1
        rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        rows = [r for r in rows if r != totals_row]

        if not rows:
            QMessageBox.information(self, "Удаление", "Выберите строку(и) в таблице для удаления.")
            return

        for r in rows:
            self.table.removeRow(r)

        self._renumber_table()
        self._setup_totals_row(preserve_values=True)
        self._adjust_table_height_to_rows()

    def _recalc_reference(self) -> None:
        spice_sum = self.sb_spice_dishes.value() * self.sp_spice_per_dish.value()
        salt_sum = self.sb_salt_dishes.value() * self.sp_salt_per_dish.value()
        total = spice_sum + salt_sum

        self.sp_spice_sum.blockSignals(True)
        self.sp_salt_sum.blockSignals(True)
        self.sp_ref_total.blockSignals(True)
        self.sp_underspend.blockSignals(True)

        self.sp_spice_sum.setValue(spice_sum)
        self.sp_salt_sum.setValue(salt_sum)
        self.sp_ref_total.setValue(total)

        underspend = max(0.0, self.sp_control_consumed.value() - total)
        self.sp_underspend.setValue(underspend)

        self.sp_spice_sum.blockSignals(False)
        self.sp_salt_sum.blockSignals(False)
        self.sp_ref_total.blockSignals(False)
        self.sp_underspend.blockSignals(False)

    def _on_compiler_name_changed(self, full_name: str) -> None:
        if self._syncing_person_fields:
            return
        self._syncing_person_fields = True
        try:
            position = PERSON_TO_POSITION.get(full_name, "")
            if position in POSITIONS_COMPILER:
                self.cb_compiler_position.setCurrentText(position)
                names = POSITION_TO_PERSONS.get(position, [])
                self._replace_combo_items(self.cb_compiler_name, names, selected_text=full_name)
        finally:
            self._syncing_person_fields = False
        self.ed_compiler_signature.setText(full_name_to_signature(full_name))

    def _on_compiler_position_changed(self, position: str) -> None:
        if self._syncing_person_fields:
            return
        self._syncing_person_fields = True
        try:
            names = POSITION_TO_PERSONS.get(position, []) if position else NAMES_COMPILER
            selected = names[0] if position and names else ""
            self._replace_combo_items(self.cb_compiler_name, names, selected_text=selected)
            self.ed_compiler_signature.setText(full_name_to_signature(self.cb_compiler_name.currentText()))
        finally:
            self._syncing_person_fields = False

    def _on_head_position_changed(self, position: str) -> None:
        if self._syncing_person_fields:
            return
        self._syncing_person_fields = True
        try:
            names = POSITION_TO_PERSONS.get(position, []) if position else NAMES_APPROVE
            selected = names[0] if position and names else ""
            self._replace_combo_items(self.cb_head_name, names, selected_text=selected)
        finally:
            self._syncing_person_fields = False

    def _on_head_name_changed(self, full_name: str) -> None:
        if self._syncing_person_fields:
            return
        self._syncing_person_fields = True
        try:
            position = PERSON_TO_POSITION.get(full_name, "")
            if position in POSITIONS_APPROVE:
                self.cb_head_position.setCurrentText(position)
                names = POSITION_TO_PERSONS.get(position, [])
                self._replace_combo_items(self.cb_head_name, names, selected_text=full_name)
        finally:
            self._syncing_person_fields = False

    def _on_accountant_name_changed(self, full_name: str) -> None:
        self.ed_accountant_signature.setText(full_name_to_signature(full_name))

    def clear_form(self) -> None:
        for ed in (
            self.ed_doc_no,
            self.ed_okpo,
            self.ed_org,
            self.ed_dept,
            self.ed_okdp,
            self.ed_operation,
            self.ed_compiler_signature,
            self.ed_accountant_signature,
        ):
            ed.clear()

        for cb in (
            self.cb_head_position,
            self.cb_head_name,
            self.cb_compiler_position,
            self.cb_compiler_name,
            self.cb_accountant_name,
        ):
            cb.setCurrentIndex(0)
            cb.setEditText("")

        today = QDate.currentDate()
        self.ed_doc_date.setDate(today)
        self._set_default_reporting_period()
        self.ed_act_date.setDate(today)
        self.ed_accountant_position.setText("Бухгалтер")
        self._initialize_people_selectors()

        current_data_rows = max(0, self.table.rowCount() - 1)
        self.table.clearSpans()
        self.table.setRowCount(current_data_rows + 1)
        self.table.clearContents()

        self._renumber_table()
        self._setup_totals_row(preserve_values=False)

        for w in (self.ed_total_open, self.ed_total_recv, self.ed_total_close, self.ed_total_cons):
            w.setValue(0.0)

        self._adjust_table_height_to_rows()

        self.sb_spice_dishes.setValue(0)
        self.sb_salt_dishes.setValue(0)
        self.sp_spice_per_dish.setValue(0.0)
        self.sp_salt_per_dish.setValue(0.0)
        self.sp_control_consumed.setValue(0.0)
        self._recalc_reference()

def main() -> int:
    app = QApplication(sys.argv)
    win = OP13BlankWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())