from __future__ import annotations

import sys
from dataclasses import dataclass
from typing import Optional

from PySide6.QtCore import Qt, QDate, QLocale, Signal
from PySide6.QtGui import QAction, QIntValidator
from PySide6.QtWidgets import (
    QApplication,
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
NAMES_APPROVE = [
    "Иванов Иван Иванович",
    "Смирнов Алексей Петрович",
    "Кузнецова Мария Сергеевна",
]

POSITIONS_COMPILER = [
    "Кладовщик",
    "Заведующий складом",
]
NAMES_COMPILER = [
    "Соколов Дмитрий Андреевич",
    "Васильева Ольга Николаевна",
]

NAMES_ACCOUNTANT = [
    "Петрова Анна Сергеевна",
    "Егорова Наталья Викторовна",
]


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


class OP13BlankWindow(QMainWindow):
    DATA_ROWS_DEFAULT = 3

    def __init__(self):
        super().__init__()
        self.setWindowTitle("ОП-13 - Акт расхода специй и соли (вариант 1)")
        self.resize(1100, 800)

        self._did_first_table_resize = False

        self._build_actions()
        self._build_menu()

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        root = QWidget()
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

    def _build_title_block(self) -> QWidget:
        box = QGroupBox("Унифицированная форма № ОП-13 - Акт расхода специй и соли")
        layout = QVBoxLayout(box)
        layout.setSpacing(10)

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(8)

        self.ed_doc_no = QLineEdit()
        self.ed_doc_no.setObjectName("doc_no")
        self.ed_doc_no.setPlaceholderText("Номер документа")

        self.ed_doc_date = self._make_date_edit("dd.MM.yyyy")
        self.ed_doc_date.setObjectName("doc_date")

        self.ed_period_from = self._make_date_edit("dd.MM.yyyy")
        self.ed_period_from.setObjectName("period_from")

        self.ed_period_to = self._make_date_edit("dd.MM.yyyy")
        self.ed_period_to.setObjectName("period_to")

        grid.addWidget(QLabel("Номер документа:"), 0, 0)
        grid.addWidget(self.ed_doc_no, 0, 1)
        grid.addWidget(QLabel("Дата составления:"), 0, 2)
        grid.addWidget(self.ed_doc_date, 0, 3)

        grid.addWidget(QLabel("Отчетный период: с"), 1, 0)
        grid.addWidget(self.ed_period_from, 1, 1)
        grid.addWidget(QLabel("по"), 1, 2, alignment=Qt.AlignmentFlag.AlignRight)
        grid.addWidget(self.ed_period_to, 1, 3)

        grid.setColumnStretch(1, 2)
        grid.setColumnStretch(3, 2)
        layout.addLayout(grid)

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

        self.ed_okud = QLineEdit("0330513")
        self.ed_okud.setObjectName("okud")
        self.ed_okud.setReadOnly(True)
        self.ed_okud.setMaximumWidth(140)

        self.ed_okpo = QLineEdit()
        self.ed_okpo.setObjectName("okpo")
        self.ed_okpo.setValidator(QIntValidator(0, 999999999))
        self.ed_okpo.setPlaceholderText("ОКПО")
        self.ed_okpo.setMaximumWidth(140)

        self.ed_okdp = QLineEdit()
        self.ed_okdp.setObjectName("okdp")
        self.ed_okdp.setPlaceholderText("Вид деятельности по ОКДП")

        self.ed_operation = QLineEdit()
        self.ed_operation.setObjectName("operation_kind")
        self.ed_operation.setPlaceholderText("Вид операции")

        approve_box = QGroupBox("УТВЕРЖДАЮ")
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
        self.cb_head_name.currentTextChanged.connect(self._on_head_name_changed)

        self.ed_head_signature = QLineEdit()
        self.ed_head_signature.setObjectName("head_signature")
        self.ed_head_signature.setPlaceholderText("Подпись")
        self.ed_head_signature.setReadOnly(True)

        self.ed_act_date = self._make_date_edit("«dd» MMMM yyyy 'г.'")
        self.ed_act_date.setObjectName("approval_date")

        approve_grid.addWidget(QLabel("Должность:"), 0, 0)
        approve_grid.addWidget(self.cb_head_position, 0, 1)
        approve_grid.addWidget(QLabel("Подпись:"), 1, 0)
        approve_grid.addWidget(self.ed_head_signature, 1, 1)
        approve_grid.addWidget(QLabel("Расшифровка:"), 2, 0)
        approve_grid.addWidget(self.cb_head_name, 2, 1)
        approve_grid.addWidget(QLabel("Дата утверждения:"), 3, 0)
        approve_grid.addWidget(self.ed_act_date, 3, 1)
        approve_box.setMinimumWidth(450)

        top_row = QHBoxLayout()
        top_row.setSpacing(12)

        left_grid = QGridLayout()
        left_grid.setHorizontalSpacing(6)
        left_grid.setVerticalSpacing(6)
        left_grid.addWidget(QLabel("Организация:"), 0, 0)
        left_grid.addWidget(self.ed_org, 0, 1, 1, 3)
        left_grid.addWidget(QLabel("Подразделение:"), 1, 0)
        left_grid.addWidget(self.ed_dept, 1, 1, 1, 3)
        left_grid.addWidget(QLabel("ОКУД:"), 2, 0)
        left_grid.addWidget(self.ed_okud, 2, 1)
        left_grid.addWidget(QLabel("ОКДП:"), 2, 2)
        left_grid.addWidget(self.ed_okdp, 2, 3)
        left_grid.addWidget(QLabel("ОКПО:"), 3, 0)
        left_grid.addWidget(self.ed_okpo, 3, 1)
        left_grid.addWidget(QLabel("Вид операции:"), 3, 2)
        left_grid.addWidget(self.ed_operation, 3, 3)
        left_grid.setColumnStretch(1, 0)
        left_grid.setColumnStretch(3, 1)

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

        hdr = self.table.horizontalHeader()
        hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)

        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(0, 50)

        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
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

        grid.addWidget(mk_label("Итого", bold=True), 3, 0)
        grid.addWidget(QLabel(""), 3, 1)
        grid.addWidget(QLabel(""), 3, 2)
        grid.addWidget(self.sp_ref_total, 3, 3)

        grid.addWidget(mk_label("Израсходовано согласно контрольного расчета:"), 4, 0, 1, 3)
        grid.addWidget(self.sp_control_consumed, 4, 3)

        grid.addWidget(mk_label("Сумма недорасхода:"), 5, 0, 1, 3)
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

        grid.addWidget(QLabel("Расчет и справку составил:"), 0, 0)
        grid.addWidget(self.cb_compiler_position, 0, 1)
        grid.addWidget(self.ed_compiler_signature, 0, 2)
        grid.addWidget(self.cb_compiler_name, 0, 3)

        grid.addWidget(QLabel("Бухгалтер:"), 1, 0)
        grid.addWidget(QLabel(""), 1, 1)
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

        row.addStretch(1)

        self.btn_print = QPushButton("Печать")
        self.btn_print.setObjectName("btn_print")
        self.btn_print.setStyleSheet(
            "QPushButton { background-color: #2563eb; color: white; border-radius: 4px; padding: 6px 16px; }"
            "QPushButton:hover { background-color: #1d4ed8; }"
            "QPushButton:pressed { background-color: #1e40af; }"
        )

        self.btn_export_xls = QPushButton("Выгрузка в XLS")
        self.btn_export_xls.setObjectName("btn_export_xls")
        self.btn_export_xls.setStyleSheet(
            "QPushButton { background-color: #16a34a; color: white; border-radius: 4px; padding: 6px 16px; }"
            "QPushButton:hover { background-color: #15803d; }"
            "QPushButton:pressed { background-color: #166534; }"
        )

        row.addWidget(self.btn_print)
        row.addWidget(self.btn_export_xls)

        return box

    def _build_actions(self) -> None:
        self.act_clear = QAction("Очистить", self)
        self.act_clear.setShortcut("Ctrl+N")
        self.act_clear.triggered.connect(self.clear_form)

        self.act_exit = QAction("Выход", self)
        self.act_exit.setShortcut("Ctrl+Q")
        self.act_exit.triggered.connect(self.close)

        self.act_about = QAction("О программе", self)
        self.act_about.triggered.connect(self._about)

    def _build_menu(self) -> None:
        m_file = self.menuBar().addMenu("Файл")
        m_file.addAction(self.act_clear)
        m_file.addSeparator()
        m_file.addAction(self.act_exit)

        m_help = self.menuBar().addMenu("Справка")
        m_help.addAction(self.act_about)

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

    def _on_head_name_changed(self, full_name: str) -> None:
        self.ed_head_signature.setText(full_name_to_signature(full_name))

    def _on_compiler_name_changed(self, full_name: str) -> None:
        self.ed_compiler_signature.setText(full_name_to_signature(full_name))

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
            self.ed_head_signature,
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
        self.ed_period_from.setDate(today)
        self.ed_period_to.setDate(today)
        self.ed_act_date.setDate(today)

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

    def _about(self) -> None:
        QMessageBox.information(
            self,
            "О программе",
            "Интерфейс для заполнения унифицированной формы № ОП-13 "
            "«Акт расхода специй и соли».\n"
            "Реализация: PySide6 / Qt.",
        )


def main() -> int:
    app = QApplication(sys.argv)
    win = OP13BlankWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())