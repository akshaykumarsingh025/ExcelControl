from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QVBoxLayout, QGroupBox, QLabel, QTextEdit, QLineEdit,
    QComboBox, QCheckBox, QSpinBox, QPushButton, QScrollArea,
    QWidget, QFormLayout,
)
from ui.workflow_base import WorkflowBase
from core.features import (
    build_lead_scoring_prompt,
    build_campaign_roi_prompt,
    build_sentiment_prompt,
    build_email_cleaner_prompt,
)


class MarketingPage(WorkflowBase):
    def get_workflow_name(self):
        return "Marketing & Sales"

    def get_workflow_description(self):
        return "Lead scoring, campaign ROI analysis, sentiment analysis, and email list cleaning."

    def setup_ui(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setSpacing(16)
        layout.setContentsMargins(8, 8, 8, 8)

        layout.addWidget(self._build_lead_scoring())
        layout.addWidget(self._build_campaign_roi())
        layout.addWidget(self._build_sentiment())
        layout.addWidget(self._build_email_cleaner())
        layout.addStretch()

        scroll.setWidget(container)
        self.get_content_layout().addWidget(scroll)

    def _build_lead_scoring(self):
        group = QGroupBox("D1: Lead Scoring Engine")
        form = QFormLayout()
        form.setSpacing(8)

        self.ls_company_size = QLineEdit("Company Size")
        form.addRow("Company Size Column:", self.ls_company_size)

        self.ls_engagement = QLineEdit("Engagement Score")
        form.addRow("Engagement Score Column:", self.ls_engagement)

        self.ls_industry = QLineEdit("Industry")
        form.addRow("Industry Column:", self.ls_industry)

        self.ls_budget = QLineEdit("Budget")
        form.addRow("Budget Column:", self.ls_budget)

        self.ls_rules = QTextEdit()
        self.ls_rules.setPlaceholderText(
            "Company Size >500 = 30pts\n"
            "Engagement Score >7 = 25pts\n"
            "Budget >50000 = 20pts\n"
            "Industry = Tech = 15pts\n"
            "Otherwise = 10pts"
        )
        self.ls_rules.setMaximumHeight(100)
        form.addRow("Scoring Rules:", self.ls_rules)

        btn = QPushButton("Score All Leads")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_lead_scoring)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_campaign_roi(self):
        group = QGroupBox("D2: Campaign ROI Calculator")
        form = QFormLayout()
        form.setSpacing(8)

        self.roi_name_col = QLineEdit("Campaign")
        form.addRow("Campaign Name Column:", self.roi_name_col)

        self.roi_spend_col = QLineEdit("Spend")
        form.addRow("Spend Column:", self.roi_spend_col)

        self.roi_clicks_col = QLineEdit("Clicks")
        form.addRow("Clicks Column:", self.roi_clicks_col)

        self.roi_conversions_col = QLineEdit("Conversions")
        form.addRow("Conversions Column:", self.roi_conversions_col)

        self.roi_revenue_col = QLineEdit("Revenue")
        form.addRow("Revenue Column:", self.roi_revenue_col)

        btn = QPushButton("Calculate Campaign ROI")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_campaign_roi)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_sentiment(self):
        group = QGroupBox("D3: Sentiment Analyzer")
        form = QFormLayout()
        form.setSpacing(8)

        self.sent_column = QLineEdit("Review")
        form.addRow("Review/Feedback Column:", self.sent_column)

        self.sent_format = QComboBox()
        self.sent_format.addItems([
            "Simple (Positive/Negative/Neutral)",
            "Detailed with Score (-1.0 to 1.0)",
            "Extract Key Themes",
        ])
        form.addRow("Output Format:", self.sent_format)

        self.sent_max_rows = QSpinBox()
        self.sent_max_rows.setRange(10, 50000)
        self.sent_max_rows.setValue(500)
        form.addRow("Max Rows to Analyze:", self.sent_max_rows)

        btn = QPushButton("Analyze Sentiment")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_sentiment)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _build_email_cleaner(self):
        group = QGroupBox("D4: Email List Cleaner")
        form = QFormLayout()
        form.setSpacing(8)

        self.ec_duplicates = QCheckBox("Remove duplicates")
        self.ec_duplicates.setChecked(True)
        form.addRow("", self.ec_duplicates)

        self.ec_case = QCheckBox("Fix capitalization (lowercase emails)")
        self.ec_case.setChecked(True)
        form.addRow("", self.ec_case)

        self.ec_validate = QCheckBox("Validate email format")
        self.ec_validate.setChecked(True)
        form.addRow("", self.ec_validate)

        self.ec_names = QCheckBox("Separate first/last name")
        self.ec_names.setChecked(False)
        form.addRow("", self.ec_names)

        self.ec_role = QCheckBox("Remove role-based emails (info@, support@, etc.)")
        self.ec_role.setChecked(False)
        form.addRow("", self.ec_role)

        btn = QPushButton("Clean Email List")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self._on_email_cleaner)
        form.addRow("", btn)

        group.setLayout(form)
        return group

    def _on_lead_scoring(self):
        prompt = build_lead_scoring_prompt(
            company_size_col=self.ls_company_size.text().strip(),
            engagement_col=self.ls_engagement.text().strip(),
            industry_col=self.ls_industry.text().strip(),
            budget_col=self.ls_budget.text().strip(),
            rules=self.ls_rules.toPlainText().strip(),
        )
        self.command_sent.emit(prompt)

    def _on_campaign_roi(self):
        prompt = build_campaign_roi_prompt(
            name_col=self.roi_name_col.text().strip(),
            spend_col=self.roi_spend_col.text().strip(),
            clicks_col=self.roi_clicks_col.text().strip(),
            conversions_col=self.roi_conversions_col.text().strip(),
            revenue_col=self.roi_revenue_col.text().strip(),
        )
        self.command_sent.emit(prompt)

    def _on_sentiment(self):
        prompt = build_sentiment_prompt(
            column_name=self.sent_column.text().strip(),
            output_format=self.sent_format.currentText(),
            max_rows=self.sent_max_rows.value(),
        )
        self.command_sent.emit(prompt)

    def _on_email_cleaner(self):
        options = {
            "remove_duplicates": self.ec_duplicates.isChecked(),
            "fix_case": self.ec_case.isChecked(),
            "remove_invalid": self.ec_validate.isChecked(),
            "trim_whitespace": True,
            "remove_unsubscribed": False,
            "categorize_domain": self.ec_names.isChecked(),
        }
        prompt = build_email_cleaner_prompt(options)
        self.command_sent.emit(prompt)
