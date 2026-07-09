from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QVBoxLayout, QGroupBox, QLabel, QLineEdit,
    QComboBox, QCheckBox, QSpinBox, QDoubleSpinBox, QPushButton,
    QScrollArea, QWidget, QFormLayout,
)
from ui.workflow_base import WorkflowBase
from core.features import (
    build_emi_prompt,
    build_property_comparison_prompt,
    build_rental_yield_prompt,
)


class RealEstatePage(WorkflowBase):
    def get_workflow_name(self):
        return "Real Estate"

    def get_workflow_description(self):
        return "EMI calculators, property comparisons, and rental yield analysis."

    def setup_ui(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setSpacing(16)
        layout.setContentsMargins(8, 8, 8, 8)

        layout.addWidget(self._build_emi())
        layout.addWidget(self._build_property_comparison())
        layout.addWidget(self._build_rental_yield())
        layout.addStretch()

        scroll.setWidget(container)
        self.get_content_layout().addWidget(scroll)

    def _build_emi(self):
        group = QGroupBox("H1: EMI Calculator")
        form = QFormLayout()
        form.setSpacing(8)

        self.emi_principal = QDoubleSpinBox()
        self.emi_principal.setRange(100000, 100000000)
        self.emi_principal.setValue(5000000)
        self.emi_principal.setPrefix("₹ ")
        self.emi_principal.setSingleStep(100000)
        form.addRow("Loan Amount:", self.emi_principal)

        self.emi_rate = QDoubleSpinBox()
        self.emi_rate.setRange(0.1, 30.0)
        self.emi_rate.setValue(8.5)
        self.emi_rate.setSuffix(" %")
        self.emi_rate.setSingleStep(0.25)
        form.addRow("Annual Interest Rate:", self.emi_rate)

        self.emi_tenure = QSpinBox()
        self.emi_tenure.setRange(1, 30)
        self.emi_tenure.setValue(20)
        self.emi_tenure.setSuffix(" years")
        form.addRow("Loan Tenure:", self.emi_tenure)

        self.emi_amortization = QCheckBox("Include amortization schedule")
        self.emi_amortization.setChecked(True)
        form.addRow("", self.emi_amortization)

        self.emi_compare = QCheckBox("Compare multiple interest rates")
        self.emi_compare.setChecked(False)
        form.addRow("", self.emi_compare)

        btn = QPushButton("Calculate EMI")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_emi)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_property_comparison(self):
        group = QGroupBox("H2: Property Comparison Sheet")
        form = QFormLayout()
        form.setSpacing(8)

        self.prop_count = QSpinBox()
        self.prop_count.setRange(2, 20)
        self.prop_count.setValue(5)
        form.addRow("Number of Properties:", self.prop_count)

        self.prop_criteria = QLineEdit()
        self.prop_criteria.setPlaceholderText(
            "Price, Location, Size (sqft), Bedrooms, Bathrooms, Parking, Amenities Score"
        )
        form.addRow("Comparison Criteria (comma-separated):", self.prop_criteria)

        self.prop_roi = QCheckBox("Include ROI estimate")
        self.prop_roi.setChecked(True)
        form.addRow("", self.prop_roi)

        btn = QPushButton("Create Comparison Sheet")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_property_comparison)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_rental_yield(self):
        group = QGroupBox("H3: Rental Yield Calculator")
        form = QFormLayout()
        form.setSpacing(8)

        self.yield_price_col = QLineEdit("Purchase Price")
        form.addRow("Purchase Price Column:", self.yield_price_col)

        self.yield_rent_col = QLineEdit("Monthly Rent")
        form.addRow("Monthly Rent Column:", self.yield_rent_col)

        self.yield_maintenance_col = QLineEdit("Annual Maintenance")
        form.addRow("Annual Maintenance Column:", self.yield_maintenance_col)

        self.yield_vacancy = QDoubleSpinBox()
        self.yield_vacancy.setRange(0, 30)
        self.yield_vacancy.setValue(5)
        self.yield_vacancy.setSuffix(" %")
        form.addRow("Vacancy Rate:", self.yield_vacancy)

        btn = QPushButton("Calculate Rental Yield")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_rental_yield)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _on_emi(self):
        prompt = build_emi_prompt(
            principal=self.emi_principal.value(),
            rate=self.emi_rate.value(),
            tenure=self.emi_tenure.value() * 12,
            include_amortization=self.emi_amortization.isChecked(),
            compare_rates=self.emi_compare.isChecked(),
        )
        self.command_sent.emit(prompt)

    def _on_property_comparison(self):
        prompt = build_property_comparison_prompt(
            num_properties=self.prop_count.value(),
            criteria=self.prop_criteria.text().strip() or "Price, Location, Size, Bedrooms",
            include_roi=self.prop_roi.isChecked(),
        )
        self.command_sent.emit(prompt)

    def _on_rental_yield(self):
        prompt = build_rental_yield_prompt(
            price_col=self.yield_price_col.text().strip(),
            rent_col=self.yield_rent_col.text().strip(),
            maintenance_col=self.yield_maintenance_col.text().strip(),
            vacancy_rate=self.yield_vacancy.value(),
        )
        self.command_sent.emit(prompt)
