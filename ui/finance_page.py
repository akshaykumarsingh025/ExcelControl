from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QVBoxLayout, QGroupBox, QLabel, QTextEdit, QLineEdit,
    QComboBox, QCheckBox, QDoubleSpinBox, QPushButton, QScrollArea,
    QWidget, QFormLayout, QFileDialog,
)
from ui.workflow_base import WorkflowBase
from core.features import (
    build_reconciliation_prompt,
    build_invoice_extract_prompt,
    build_pl_variance_prompt,
    build_tax_classifier_prompt,
)


class FinancePage(WorkflowBase):
    def get_workflow_name(self):
        return "Finance & Accounting"

    def get_workflow_description(self):
        return "Bank reconciliation, invoice extraction, P&L statements, budget variance, and tax classification."

    def setup_ui(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setSpacing(16)
        layout.setContentsMargins(8, 8, 8, 8)

        layout.addWidget(self._build_bank_reconciliation())
        layout.addWidget(self._build_invoice_extractor())
        layout.addWidget(self._build_pl_builder())
        layout.addWidget(self._build_budget_variance())
        layout.addWidget(self._build_tax_classifier())
        layout.addStretch()

        scroll.setWidget(container)
        self.get_content_layout().addWidget(scroll)

    def _get_sheet_names(self):
        try:
            if hasattr(self, 'excel') and self.excel and self.excel.is_connected():
                return self.excel.get_sheet_names()
        except Exception:
            pass
        return ["Sheet1"]

    def _build_bank_reconciliation(self):
        group = QGroupBox("B1: Bank Reconciliation")
        form = QFormLayout()
        form.setSpacing(8)

        self.recon_bank_sheet = QComboBox()
        self.recon_bank_sheet.setEditable(True)
        form.addRow("Bank Statement Sheet:", self.recon_bank_sheet)

        self.recon_ledger_sheet = QComboBox()
        self.recon_ledger_sheet.setEditable(True)
        form.addRow("Ledger Sheet:", self.recon_ledger_sheet)

        self.recon_tolerance = QLineEdit("0.01")
        form.addRow("Tolerance Amount:", self.recon_tolerance)

        btn = QPushButton("Run Reconciliation")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_reconciliation)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_invoice_extractor(self):
        group = QGroupBox("B2: Invoice Extractor")
        form = QFormLayout()
        form.setSpacing(8)

        self.invoice_path_label = QLabel("No file selected")
        self.invoice_path_label.setObjectName("subheadingLabel")
        form.addRow("Selected File:", self.invoice_path_label)

        self.invoice_format = QComboBox()
        self.invoice_format.addItems(["Standard", "GST", "Proforma", "Receipt", "Custom"])
        form.addRow("Invoice Format:", self.invoice_format)

        btn_select = QPushButton("Select Invoice Image/PDF")
        btn_select.clicked.connect(self._on_select_invoice_file)
        form.addRow("", btn_select)

        btn = QPushButton("Extract Invoice Data")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_extract_invoice)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_pl_builder(self):
        group = QGroupBox("B3: P&L Statement Builder")
        form = QFormLayout()
        form.setSpacing(8)

        self.pl_sheet = QComboBox()
        self.pl_sheet.setEditable(True)
        form.addRow("Transaction Data Sheet:", self.pl_sheet)

        self.pl_revenue_col = QLineEdit("Amount")
        form.addRow("Revenue Column:", self.pl_revenue_col)

        self.pl_category_col = QLineEdit("Category")
        form.addRow("Category Column:", self.pl_category_col)

        self.pl_date_col = QLineEdit("Date")
        form.addRow("Date Column:", self.pl_date_col)

        self.pl_period = QComboBox()
        self.pl_period.addItems(["Monthly", "Quarterly", "Yearly"])
        form.addRow("Period:", self.pl_period)

        btn = QPushButton("Build P&L Statement")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_build_pl)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_budget_variance(self):
        group = QGroupBox("B4: Budget Variance Analyzer")
        form = QFormLayout()
        form.setSpacing(8)

        self.bv_budget_sheet = QComboBox()
        self.bv_budget_sheet.setEditable(True)
        form.addRow("Budget Sheet:", self.bv_budget_sheet)

        self.bv_actual_sheet = QComboBox()
        self.bv_actual_sheet.setEditable(True)
        form.addRow("Actual Sheet:", self.bv_actual_sheet)

        self.bv_threshold = QDoubleSpinBox()
        self.bv_threshold.setRange(0.0, 100.0)
        self.bv_threshold.setValue(10.0)
        self.bv_threshold.setSuffix(" %")
        form.addRow("Variance Threshold:", self.bv_threshold)

        btn = QPushButton("Analyze Budget Variance")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_budget_variance)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_tax_classifier(self):
        group = QGroupBox("B5: Tax Category Classifier")
        form = QFormLayout()
        form.setSpacing(8)

        self.tax_expense_col = QLineEdit("Expense")
        form.addRow("Expense Column Name:", self.tax_expense_col)

        self.tax_gst = QCheckBox("Include Indian GST categories")
        self.tax_gst.setChecked(True)
        form.addRow("", self.tax_gst)

        self.tax_us = QCheckBox("Include US tax categories")
        self.tax_us.setChecked(False)
        form.addRow("", self.tax_us)

        btn = QPushButton("Classify Expenses")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_tax_classify)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _populate_sheet_combos(self):
        sheets = self._get_sheet_names()
        for combo in [self.recon_bank_sheet, self.recon_ledger_sheet,
                      self.pl_sheet, self.bv_budget_sheet, self.bv_actual_sheet]:
            combo.clear()
            combo.addItems(sheets)

    def showEvent(self, event):
        super().showEvent(event)
        self._populate_sheet_combos()

    def _on_select_invoice_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Invoice", "",
            "Images & PDFs (*.png *.jpg *.jpeg *.bmp *.pdf *.tiff);;All Files (*)"
        )
        if path:
            self.invoice_path_label.setText(path)

    def _on_reconciliation(self):
        prompt = build_reconciliation_prompt(
            bank_sheet=self.recon_bank_sheet.currentText(),
            ledger_sheet=self.recon_ledger_sheet.currentText(),
            tolerance=self.recon_tolerance.text().strip(),
        )
        self.command_sent.emit(prompt)

    def _on_extract_invoice(self):
        path = self.invoice_path_label.text()
        fmt = self.invoice_format.currentText()
        if path and path != "No file selected":
            prompt = (
                f"Extract data from this invoice image and write it to Excel.\n"
                f"Invoice format: {fmt}\n\n"
                f"Use the image at: {path}\n"
                "Read the image, extract all invoice fields, and create a formatted invoice sheet.\n"
            ) + build_invoice_extract_prompt(fmt)
        else:
            prompt = build_invoice_extract_prompt(fmt)
        self.command_sent.emit(prompt)

    def _on_build_pl(self):
        prompt = build_pl_variance_prompt(
            budget_sheet=f"Budget ({self.pl_sheet.currentText()})",
            actual_sheet=f"Actual ({self.pl_sheet.currentText()})",
            threshold=self.bv_threshold.value(),
        )
        self.command_sent.emit(prompt)

    def _on_budget_variance(self):
        prompt = build_pl_variance_prompt(
            budget_sheet=self.bv_budget_sheet.currentText(),
            actual_sheet=self.bv_actual_sheet.currentText(),
            threshold=self.bv_threshold.value(),
        )
        self.command_sent.emit(prompt)

    def _on_tax_classify(self):
        prompt = build_tax_classifier_prompt(
            expense_col=self.tax_expense_col.text().strip(),
            include_gst=self.tax_gst.isChecked(),
            include_us=self.tax_us.isChecked(),
        )
        self.command_sent.emit(prompt)
