from __future__ import annotations

import sys
from typing import Optional

from PySide6.QtCore import Qt, QDate, QEvent, QLocale, QObject, Signal
from PySide6.QtGui import QAction, QIntValidator, QPalette
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDateEdit,
    QDoubleSpinBox,
    QFormLayout,
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
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


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
    initials = "".join(f"{p[0]}." for p in parts[1:] if p)
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


class _PaletteChangeFilter(QObject):
    def __init__(self, callback, parent: Optional[QObject] = None):
        super().__init__(parent)
        self._callback = callback

    def eventFilter(self, watched, event):
        if event.type() == QEvent.Type.ApplicationPaletteChange:
            self._callback()
        return False


class OP13WizardWindow(QMainWindow):
    DATA_ROWS_DEFAULT = 3

    def __init__(self):
        super().__init__()
        self.setWindowTitle("ОП-13 - Акт расхода специй и соли (вариант 2)")
        self.resize(1100, 820)

        self._did_first_table_resize = False
        self._step_titles = [
            "Шаг 1. Документ и реквизиты",
            "Шаг 2. Таблица расхода",
            "Шаг 3. Справка о стоимости",
            "Шаг 4. Подписи",
        ]
        self._step_buttons: list[QPushButton] = []

        self._build_actions()
        self._build_menu()
        self._build_ui()
        self._palette_filter = _PaletteChangeFilter(self._update_theme_dependent_styles, self)
        app = QApplication.instance()
        if app is not None:
            app.installEventFilter(self._palette_filter)
        self._apply_ru_locale()
        self._recalc_reference()
        self._go_to_step(0)

    def showEvent(self, event):
        super().showEvent(event)
        if not self._did_first_table_resize:
            self._did_first_table_resize = True
            self._adjust_table_height_to_rows()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(12)

        header_box = QGroupBox("Заполнение формы по шагам")
        header_layout = QVBoxLayout(header_box)
        header_layout.setSpacing(10)

        self.lb_step = QLabel()
        self._update_theme_dependent_styles()
        header_layout.addWidget(self.lb_step)

        steps_row = QHBoxLayout()
        steps_row.setSpacing(8)
        for index, title in enumerate(self._step_titles):
            btn = QPushButton(title)
            btn.setCheckable(True)
            btn.clicked.connect(lambda checked=False, i=index: self._go_to_step(i))
            btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            self._step_buttons.append(btn)
            steps_row.addWidget(btn)
        header_layout.addLayout(steps_row)
        root.addWidget(header_box)

        self.pages = QStackedWidget()
        self.pages.addWidget(self._wrap_page(self._build_step1_page()))
        self.pages.addWidget(self._wrap_page(self._build_step2_page()))
        self.pages.addWidget(self._wrap_page(self._build_step3_page()))
        self.pages.addWidget(self._wrap_page(self._build_step4_page()))
        root.addWidget(self.pages, 1)

        nav_row = QHBoxLayout()
        nav_row.setSpacing(8)

        self.btn_prev = QPushButton("← Назад")
        self.btn_prev.clicked.connect(self._prev_step)

        self.btn_next = QPushButton("Далее →")
        self.btn_next.clicked.connect(self._next_step)

        nav_row.addWidget(self.btn_prev)
        nav_row.addWidget(self.btn_next)
        nav_row.addStretch(1)
        nav_row.addWidget(self._build_bottom_buttons_block())
        root.addLayout(nav_row)

    def _wrap_page(self, content: QWidget) -> QWidget:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        layout.addWidget(content)
        layout.addStretch(1)

        scroll.setWidget(container)
        return scroll

    def _build_step1_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setSpacing(12)
        layout.addWidget(self._build_title_block())
        layout.addWidget(self._build_requisites_block())
        return page

    def _build_step2_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setSpacing(12)
        layout.addWidget(self._build_table_block())
        return page

    def _build_step3_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setSpacing(12)
        layout.addWidget(self._build_reference_block())
        return page

    def _build_step4_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setSpacing(12)
        layout.addWidget(self._build_signatures_block())

        note = QGroupBox("Проверка перед завершением")
        note_layout = QVBoxLayout(note)
        note_layout.addWidget(QLabel(
            "Убедитесь, что заполнены все поля формы. "
            "Навигация по шагам доступна без потери введенных данных."
        ))
        layout.addWidget(note)
        return page

    def _is_dark_theme(self) -> bool:
        window_color = self.palette().color(QPalette.ColorRole.Window)
        return window_color.lightness() < 128

    def _update_theme_dependent_styles(self) -> None:
        text_color = self.palette().color(QPalette.ColorRole.WindowText).name()
        self.lb_step.setStyleSheet(f"font-size: 16px; font-weight: 600; color: {text_color};")
        if hasattr(self, "pages") and self._step_buttons:
            self._go_to_step(self.pages.currentIndex())

    def _go_to_step(self, index: int) -> None:
        index = max(0, min(index, self.pages.count() - 1))
        self.pages.setCurrentIndex(index)
        self.lb_step.setText(f"{self._step_titles[index]} ({index + 1} из {self.pages.count()})")

        for i, btn in enumerate(self._step_buttons):
            btn.blockSignals(True)
            btn.setChecked(i == index)
            btn.blockSignals(False)

        default_text_color = self.palette().color(QPalette.ColorRole.ButtonText).name()
        if self._is_dark_theme():
            active_style = (
                "QPushButton {"
                "background-color: #2d4a6e;"
                "border: 1px solid #6f96bf;"
                "color: #e8f4fc;"
                "font-weight: 600;"
                "padding: 8px;"
                "}"
            )
        else:
            active_style = (
                "QPushButton {"
                "background-color: #d9ecff;"
                "border: 1px solid #7aa7d9;"
                "color: #1a1a2e;"
                "font-weight: 600;"
                "padding: 8px;"
                "}"
            )
        default_style = f"QPushButton {{ padding: 8px; color: {default_text_color}; }}"
        for i, btn in enumerate(self._step_buttons):
            btn.setStyleSheet(active_style if i == index else default_style)

        self.btn_prev.setEnabled(index > 0)
        self.btn_next.setEnabled(index < self.pages.count() - 1)
        self.btn_next.setText("Далее →" if index < self.pages.count() - 1 else "Готово")

    def _next_step(self) -> None:
        current = self.pages.currentIndex()
        if current < self.pages.count() - 1:
            self._go_to_step(current + 1)
        else:
            QMessageBox.information(
                self,
                "Заполнение завершено",
                "Все шаги заполнены. Проверьте данные и при необходимости вернитесь к нужному шагу.",
            )

    def _prev_step(self) -> None:
        self._go_to_step(self.pages.currentIndex() - 1)

    def _make_date_edit(self, display_format: str) -> QDateEdit:
        ed = QDateEdit()
        ed.setCalendarPopup(True)
        ed.setDisplayFormat(display_format)
        ed.setDate(QDate.currentDate())
        return ed

    def _build_title_block(self) -> QWidget:
        box = QGroupBox("Унифицированная форма № ОП-13 - Акт расхода специй и соли")
        layout = QGridLayout(box)
        layout.setHorizontalSpacing(10)
        layout.setVerticalSpacing(8)

        self.ed_doc_no = QLineEdit()
        self.ed_doc_no.setPlaceholderText("Номер документа")

        self.ed_doc_date = self._make_date_edit("dd.MM.yyyy")
        self.ed_period_from = self._make_date_edit("dd.MM.yyyy")
        self.ed_period_to = self._make_date_edit("dd.MM.yyyy")

        layout.addWidget(QLabel("Номер документа:"), 0, 0)
        layout.addWidget(self.ed_doc_no, 0, 1)
        layout.addWidget(QLabel("Дата составления:"), 0, 2)
        layout.addWidget(self.ed_doc_date, 0, 3)

        layout.addWidget(QLabel("Отчетный период: с"), 1, 0)
        layout.addWidget(self.ed_period_from, 1, 1)
        layout.addWidget(QLabel("по"), 1, 2, alignment=Qt.AlignmentFlag.AlignRight)
        layout.addWidget(self.ed_period_to, 1, 3)

        layout.setColumnStretch(1, 2)
        layout.setColumnStretch(3, 2)
        return box

    def _build_requisites_block(self) -> QWidget:
        box = QGroupBox("Реквизиты")
        root = QVBoxLayout(box)
        root.setSpacing(10)

        self.ed_org = QLineEdit()
        self.ed_org.setPlaceholderText("Наименование организации")

        self.ed_dept = QLineEdit()
        self.ed_dept.setPlaceholderText("Структурное подразделение")

        self.ed_okud = QLineEdit("0330513")
        self.ed_okud.setReadOnly(True)
        self.ed_okud.setMaximumWidth(140)

        self.ed_okpo = QLineEdit()
        self.ed_okpo.setValidator(QIntValidator(0, 999999999))
        self.ed_okpo.setPlaceholderText("ОКПО")
        self.ed_okpo.setMaximumWidth(140)

        self.ed_okdp = QLineEdit()
        self.ed_okdp.setPlaceholderText("Вид деятельности по ОКДП")

        self.ed_operation = QLineEdit()
        self.ed_operation.setPlaceholderText("Вид операции")

        company_box = QGroupBox("Организация")
        company_form = QFormLayout(company_box)
        company_form.addRow("Организация:", self.ed_org)
        company_form.addRow("Подразделение:", self.ed_dept)

        codes_box = QGroupBox("Коды и операция")
        codes_form = QFormLayout(codes_box)
        codes_form.addRow("ОКУД:", self.ed_okud)
        codes_form.addRow("ОКПО:", self.ed_okpo)
        codes_form.addRow("ОКДП:", self.ed_okdp)
        codes_form.addRow("Вид операции:", self.ed_operation)

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
        self.ed_head_signature.setPlaceholderText("Подпись")
        self.ed_head_signature.setReadOnly(True)

        self.ed_act_date = self._make_date_edit("«dd» MMMM yyyy 'г.'")

        approve_box = QGroupBox("УТВЕРЖДАЮ")
        approve_form = QFormLayout(approve_box)
        approve_form.addRow("Должность:", self.cb_head_position)
        approve_form.addRow("Подпись:", self.ed_head_signature)
        approve_form.addRow("Расшифровка:", self.cb_head_name)
        approve_form.addRow("Дата утверждения:", self.ed_act_date)

        top_row = QHBoxLayout()
        top_row.setSpacing(12)
        top_row.addWidget(company_box, 1)
        top_row.addWidget(codes_box, 1)
        top_row.addWidget(approve_box, 1)
        root.addLayout(top_row)
        return box

    def _adjust_table_height_to_rows(self) -> None:
        hh = self.table.horizontalHeader()
        header_h = max(hh.height(), hh.sizeHint().height())

        height = header_h
        for row in range(self.table.rowCount()):
            height += self.table.rowHeight(row)

        height += 2 * self.table.frameWidth()
        self.table.setFixedHeight(height)

    def _create_editable_item(self, alignment: Qt.AlignmentFlag | None = None) -> QTableWidgetItem:
        item = QTableWidgetItem("")
        if alignment is not None:
            item.setTextAlignment(alignment)
        return item

    def _ensure_data_row_items(self, row: int) -> None:
        if self.table.item(row, 0) is None:
            item = QTableWidgetItem(str(row + 1))
            item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
            self.table.setItem(row, 0, item)

        if self.table.item(row, 1) is None:
            self.table.setItem(row, 1, self._create_editable_item())
        if self.table.item(row, 2) is None:
            self.table.setItem(
                row,
                2,
                self._create_editable_item(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter),
            )

        for col in range(3, 7):
            if self.table.item(row, col) is None:
                self.table.setItem(
                    row,
                    col,
                    self._create_editable_item(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter),
                )

    def _renumber_table(self) -> None:
        data_rows = self.table.rowCount() - 1
        for row in range(data_rows):
            self._ensure_data_row_items(row)
            number_item = self.table.item(row, 0)
            number_item.setText(str(row + 1))

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
        self.ed_total_recv = mk_total_spin()
        self.ed_total_close = mk_total_spin()
        self.ed_total_cons = mk_total_spin()

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
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        self._renumber_table()
        self._setup_totals_row(preserve_values=False)

        hdr = self.table.horizontalHeader()
        hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        for col in range(3, 7):
            hdr.setSectionResizeMode(col, QHeaderView.ResizeMode.Stretch)
        self.table.setColumnWidth(0, 50)

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

    def _add_table_row(self) -> None:
        insert_at = self.table.rowCount() - 1
        self.table.insertRow(insert_at)
        self._ensure_data_row_items(insert_at)
        self._renumber_table()
        self._setup_totals_row(preserve_values=True)
        self._adjust_table_height_to_rows()

    def _delete_selected_rows(self) -> None:
        totals_row = self.table.rowCount() - 1
        rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        rows = [row for row in rows if row != totals_row]

        if not rows:
            QMessageBox.information(self, "Удаление", "Выберите строку(и) в таблице для удаления.")
            return

        for row in rows:
            self.table.removeRow(row)

        self._renumber_table()
        self._setup_totals_row(preserve_values=True)
        self._adjust_table_height_to_rows()

    def _build_reference_block(self) -> QWidget:
        box = QGroupBox("Справка о стоимости специй и соли, включенной в калькуляцию блюд")
        grid = QGridLayout(box)
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(8)

        for col in range(4):
            grid.setColumnStretch(col, 1)

        def mk_label(text: str, bold: bool = False) -> QLabel:
            label = QLabel(f"<b>{text}</b>" if bold else text)
            label.setWordWrap(True)
            return label

        def expand_fixed(widget: QWidget) -> QWidget:
            widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            return widget

        self.sb_spice_dishes = QSpinBox()
        self.sb_spice_dishes.setRange(0, 10_000_000)

        self.sp_spice_per_dish = expand_fixed(MoneySpin(maximum=1_000_000))
        self.sp_spice_sum = expand_fixed(MoneySpin(maximum=1_000_000_000))
        self.sp_spice_sum.setReadOnly(True)
        self.sp_spice_sum.setButtonSymbols(QDoubleSpinBox.ButtonSymbols.NoButtons)

        self.sb_salt_dishes = QSpinBox()
        self.sb_salt_dishes.setRange(0, 10_000_000)

        self.sp_salt_per_dish = expand_fixed(MoneySpin(maximum=1_000_000))
        self.sp_salt_sum = expand_fixed(MoneySpin(maximum=1_000_000_000))
        self.sp_salt_sum.setReadOnly(True)
        self.sp_salt_sum.setButtonSymbols(QDoubleSpinBox.ButtonSymbols.NoButtons)

        self.sp_ref_total = expand_fixed(MoneySpin(maximum=1_000_000_000))
        self.sp_ref_total.setReadOnly(True)
        self.sp_ref_total.setButtonSymbols(QDoubleSpinBox.ButtonSymbols.NoButtons)

        self.sp_control_consumed = expand_fixed(MoneySpin(maximum=1_000_000_000))

        self.sp_underspend = expand_fixed(MoneySpin(maximum=1_000_000_000))
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
        self.sp_underspend.setValue(max(0.0, self.sp_control_consumed.value() - total))

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
        menu_file = self.menuBar().addMenu("Файл")
        menu_file.addAction(self.act_clear)
        menu_file.addSeparator()
        menu_file.addAction(self.act_exit)

        menu_help = self.menuBar().addMenu("Справка")
        menu_help.addAction(self.act_about)

    def _apply_ru_locale(self) -> None:
        ru = QLocale(QLocale.Language.Russian, QLocale.Country.Russia)
        QLocale.setDefault(ru)
        for widget in (self.ed_doc_date, self.ed_period_from, self.ed_period_to, self.ed_act_date):
            widget.setLocale(ru)

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
            cb.setCurrentText("")

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
        self._go_to_step(0)

    def _about(self) -> None:
        QMessageBox.information(
            self,
            "О программе",
            "Пошаговый интерфейс для заполнения унифицированной формы № ОП-13.\n"
            "Реализация: PySide6 / Qt.",
        )


def main() -> int:
    app = QApplication(sys.argv)
    win = OP13WizardWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())