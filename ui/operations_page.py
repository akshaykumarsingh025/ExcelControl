from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QVBoxLayout, QGroupBox, QLabel, QLineEdit,
    QComboBox, QCheckBox, QSpinBox, QDoubleSpinBox, QPushButton,
    QScrollArea, QWidget, QFormLayout, QFileDialog,
)
from ui.workflow_base import WorkflowBase
from core.features import (
    build_inventory_dashboard_prompt,
    build_reorder_point_prompt,
    build_shipping_tracker_prompt,
    build_consolidator_prompt,
)


class OperationsPage(WorkflowBase):
    def get_workflow_name(self):
        return "Operations & Logistics"

    def get_workflow_description(self):
        return "Inventory dashboards, reorder points, shipping trackers, and multi-file consolidation."

    def setup_ui(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setSpacing(16)
        layout.setContentsMargins(8, 8, 8, 8)

        layout.addWidget(self._build_inventory())
        layout.addWidget(self._build_reorder())
        layout.addWidget(self._build_shipping())
        layout.addWidget(self._build_consolidator())
        layout.addStretch()

        scroll.setWidget(container)
        self.get_content_layout().addWidget(scroll)

    def _build_inventory(self):
        group = QGroupBox("G1: Inventory Dashboard")
        form = QFormLayout()
        form.setSpacing(8)

        self.inv_product_col = QLineEdit("Product")
        form.addRow("Product Name Column:", self.inv_product_col)

        self.inv_stock_col = QLineEdit("Stock")
        form.addRow("Stock Column:", self.inv_stock_col)

        self.inv_reorder_col = QLineEdit("Reorder Level")
        form.addRow("Reorder Level Column:", self.inv_reorder_col)

        self.inv_price_col = QLineEdit("Unit Price")
        form.addRow("Unit Price Column:", self.inv_price_col)

        self.inv_abc = QCheckBox("Include ABC Analysis")
        self.inv_abc.setChecked(True)
        form.addRow("", self.inv_abc)

        self.inv_alerts = QCheckBox("Low stock alerts (conditional formatting)")
        self.inv_alerts.setChecked(True)
        form.addRow("", self.inv_alerts)

        btn = QPushButton("Create Inventory Dashboard")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_inventory)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_reorder(self):
        group = QGroupBox("G2: Reorder Point Calculator")
        form = QFormLayout()
        form.setSpacing(8)

        self.reorder_product_col = QLineEdit("Product")
        form.addRow("Product Column:", self.reorder_product_col)

        self.reorder_demand_col = QLineEdit("Daily Demand")
        form.addRow("Daily Demand Column:", self.reorder_demand_col)

        self.reorder_lead_col = QLineEdit("Lead Time")
        form.addRow("Lead Time Column (days):", self.reorder_lead_col)

        self.reorder_service = QDoubleSpinBox()
        self.reorder_service.setRange(80, 99.9)
        self.reorder_service.setValue(95.0)
        self.reorder_service.setSuffix(" %")
        form.addRow("Service Level:", self.reorder_service)

        self.reorder_safety = QDoubleSpinBox()
        self.reorder_safety.setRange(0.5, 3.0)
        self.reorder_safety.setValue(1.65)
        self.reorder_safety.setSingleStep(0.05)
        form.addRow("Safety Stock Factor (Z-score):", self.reorder_safety)

        btn = QPushButton("Calculate Reorder Points")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_reorder)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_shipping(self):
        group = QGroupBox("G3: Shipping Tracker")
        form = QFormLayout()
        form.setSpacing(8)

        self.ship_count = QSpinBox()
        self.ship_count.setRange(5, 500)
        self.ship_count.setValue(20)
        form.addRow("Number of Shipments:", self.ship_count)

        self.ship_carrier = QCheckBox("Include carrier dropdown")
        self.ship_carrier.setChecked(True)
        form.addRow("", self.ship_carrier)

        self.ship_eta = QCheckBox("Auto-calculate ETA")
        self.ship_eta.setChecked(True)
        form.addRow("", self.ship_eta)

        self.ship_status = QCheckBox("Status conditional formatting")
        self.ship_status.setChecked(True)
        form.addRow("", self.ship_status)

        btn = QPushButton("Create Shipping Tracker")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_shipping)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_consolidator(self):
        group = QGroupBox("G4: Multi-File Consolidator")
        form = QFormLayout()
        form.setSpacing(8)

        self.consolidate_label = QLabel("No files selected")
        self.consolidate_label.setObjectName("subheadingLabel")
        form.addRow("Selected Files:", self.consolidate_label)

        btn_select = QPushButton("Select Files to Merge")
        btn_select.clicked.connect(self._on_select_consolidate_files)
        form.addRow("", btn_select)

        self.consolidate_dupes = QCheckBox("Remove duplicates after merge")
        self.consolidate_dupes.setChecked(True)
        form.addRow("", self.consolidate_dupes)

        self.consolidate_source = QCheckBox("Add source file column")
        self.consolidate_source.setChecked(True)
        form.addRow("", self.consolidate_source)

        self.consolidate_align = QCheckBox("Align columns by header name")
        self.consolidate_align.setChecked(True)
        form.addRow("", self.consolidate_align)

        btn = QPushButton("Merge Files")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_consolidate)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _on_select_consolidate_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Excel Files to Merge", "",
            "Excel Files (*.xlsx *.xls *.xlsm);;All Files (*)"
        )
        if files:
            self.consolidate_label.setText(f"{len(files)} file(s) selected")
            self._consolidate_files = files

    def _on_inventory(self):
        prompt = build_inventory_dashboard_prompt(
            product_col=self.inv_product_col.text().strip(),
            stock_col=self.inv_stock_col.text().strip(),
            reorder_col=self.inv_reorder_col.text().strip(),
            price_col=self.inv_price_col.text().strip(),
            include_abc=self.inv_abc.isChecked(),
            low_stock_alerts=self.inv_alerts.isChecked(),
        )
        self.command_sent.emit(prompt)

    def _on_reorder(self):
        prompt = build_reorder_point_prompt(
            product_col=self.reorder_product_col.text().strip(),
            demand_col=self.reorder_demand_col.text().strip(),
            lead_time_col=self.reorder_lead_col.text().strip(),
            service_level=self.reorder_service.value(),
            safety_factor=self.reorder_safety.value(),
        )
        self.command_sent.emit(prompt)

    def _on_shipping(self):
        prompt = build_shipping_tracker_prompt(
            num_shipments=self.ship_count.value(),
            carrier_dropdown=self.ship_carrier.isChecked(),
            auto_eta=self.ship_eta.isChecked(),
            status_format="Conditional" if self.ship_status.isChecked() else "Plain",
        )
        self.command_sent.emit(prompt)

    def _on_consolidate(self):
        file_count = 0
        if hasattr(self, '_consolidate_files'):
            file_count = len(self._consolidate_files)
        prompt = build_consolidator_prompt(
            file_count=file_count or 2,
            remove_dupes=self.consolidate_dupes.isChecked(),
            add_source=self.consolidate_source.isChecked(),
            align_columns=self.consolidate_align.isChecked(),
        )
        self.command_sent.emit(prompt)
