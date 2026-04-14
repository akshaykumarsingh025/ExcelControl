# templates.py

TEMPLATES = {
    "Monthly Financial Report": {
        "description": "Create a monthly financial report with revenue, expenses, and net profit with borders and currency formatting",
        "code": (
            'ws.range("A1:D1").value = ["Category", "Budget", "Actual", "Variance"]\n'
            'ws.range("A1:D1").font.bold = True\n'
            'ws.range("A1:D1").color = (41, 128, 185)\n'
            'ws.range("A1:D1").font.color = (255, 255, 255)\n'
            'ws.range("A1:D1").api.HorizontalAlignment = -4108\n'
            'ws.range("A2:A6").value = [["Revenue"], ["COGS"], ["Gross Profit"], ["Expenses"], ["Net Profit"]]\n'
            'ws.range("B2:B6").value = [[50000], [20000], [30000], [15000], [15000]]\n'
            'ws.range("C2:C6").value = [[55000], [18000], [37000], [14000], [23000]]\n'
            'ws.range("D2:D6").formula = ["=C2-B2", "=C3-B3", "=C4-B4", "=C5-B5", "=C6-B6"]\n'
            'ws.range("D2:D6").number_format = "$#,##0"\n'
            'ws.range("B2:C6").number_format = "$#,##0"\n'
            "for i in range(1, 7):\n"
            '    ws.range("A1:D6").api.Borders(i).LineStyle = 1\n'
            '    ws.range("A1:D6").api.Borders(i).Weight = 2\n'
            'ws.range("A1:D1").column_width = 18\n'
        ),
    },
    "Invoice Generator": {
        "description": "Create a professional invoice template with borders and totals",
        "code": (
            'ws.range("A1").value = "INVOICE"\n'
            'ws.range("A1").font.size = 20\n'
            'ws.range("A1").font.bold = True\n'
            'ws.range("A3:B4").value = [["Bill To:", "Client Name"], ["Date:", "2025-01-15"]]\n'
            'ws.range("A3:A4").font.bold = True\n'
            'ws.range("A6:D6").value = ["Item", "Quantity", "Unit Price", "Total"]\n'
            'ws.range("A6:D6").font.bold = True\n'
            'ws.range("A6:D6").color = (52, 73, 94)\n'
            'ws.range("A6:D6").font.color = (255, 255, 255)\n'
            'ws.range("A6:D6").api.HorizontalAlignment = -4108\n'
            'ws.range("A7:D9").value = [["Service A", 2, 500, "=B7*C7"], ["Service B", 1, 750, "=B8*C8"], ["Service C", 3, 200, "=B9*C9"]]\n'
            'ws.range("D7:D9").number_format = "$#,##0.00"\n'
            'ws.range("C7:C9").number_format = "$#,##0.00"\n'
            "for i in range(1, 7):\n"
            '    ws.range("A6:D9").api.Borders(i).LineStyle = 1\n'
            'ws.range("A11:B11").value = ["Subtotal:", "=SUM(D7:D9)"]\n'
            'ws.range("A12:B12").value = ["Tax (18%):", "=D11*0.18"]\n'
            'ws.range("A13:B13").value = ["Total:", "=D11+D12"]\n'
            'ws.range("B11:B13").number_format = "$#,##0.00"\n'
            'ws.range("A13:B13").font.bold = True\n'
            'ws.range("A13:B13").font.size = 12\n'
            'ws.range("A1:D1").column_width = 18\n'
        ),
    },
    "Attendance Tracker": {
        "description": "Track daily attendance with conditional formatting for absent days",
        "code": (
            'ws.range("A1:E1").value = ["Employee", "Mon", "Tue", "Wed", "Thu"]\n'
            'ws.range("A1:E1").font.bold = True\n'
            'ws.range("A1:E1").color = (39, 174, 96)\n'
            'ws.range("A1:E1").font.color = (255, 255, 255)\n'
            'ws.range("A2:E5").value = [["Alice", "P", "P", "A", "P"], ["Bob", "P", "A", "P", "P"], ["Carol", "P", "P", "P", "P"], ["Dave", "A", "P", "P", "P"]]\n'
            'ws.range("F1").value = "Present Days"\n'
            'ws.range("F1").font.bold = True\n'
            'ws.range("F2:F5").formula = ["=COUNTIF(B2:E2,\\"P\\")", "=COUNTIF(B3:E3,\\"P\\")", "=COUNTIF(B4:E4,\\"P\\")", "=COUNTIF(B5:E5,\\"P\\")"]\n'
            'ws.range("B2:E5").api.FormatConditions.Add(Type=1, Operator=3, Formula1="\\"A\\"")\n'
            'ws.range("B2:E5").api.FormatConditions(1).Interior.Color = 255\n'
            'ws.range("A1:F1").column_width = 15\n'
        ),
    },
    "Project Timeline": {
        "description": "Create a project timeline with data validation for status and freeze panes",
        "code": (
            'ws.range("A1:E1").value = ["Task", "Start Date", "End Date", "Status", "Owner"]\n'
            'ws.range("A1:E1").font.bold = True\n'
            'ws.range("A1:E1").color = (142, 68, 173)\n'
            'ws.range("A1:E1").font.color = (255, 255, 255)\n'
            'ws.range("A2:E5").value = [["Planning", "2025-01-01", "2025-01-15", "Complete", "Alice"], ["Design", "2025-01-10", "2025-02-01", "In Progress", "Bob"], ["Development", "2025-02-01", "2025-03-15", "Not Started", "Carol"], ["Testing", "2025-03-10", "2025-04-01", "Not Started", "Dave"]]\n'
            'ws.range("D2:D5").api.Validation.Add(Type=3, Formula1="Complete,In Progress,Not Started,On Hold")\n'
            'ws.range("D2:D5").api.Validation.InCellDropdown = True\n'
            'ws.range("B2:C5").number_format = "YYYY-MM-DD"\n'
            'ws.freeze_panes.freeze_at("A2")\n'
            'ws.range("A1:E1").column_width = 18\n'
        ),
    },
    "Sales Dashboard Data": {
        "description": "Set up sales data as an Excel Table with chart ready to go",
        "code": (
            'ws.range("A1:C1").value = ["Month", "Sales", "Expenses"]\n'
            'ws.range("A2:C13").value = [["Jan", 12000, 8000], ["Feb", 15000, 9000], ["Mar", 18000, 11000], ["Apr", 16000, 10000], ["May", 20000, 12000], ["Jun", 22000, 13000], ["Jul", 19000, 11500], ["Aug", 25000, 14000], ["Sep", 23000, 13500], ["Oct", 27000, 15000], ["Nov", 30000, 16000], ["Dec", 35000, 18000]]\n'
            'ws.tables.add(source=ws.range("A1:C13"), name="SalesTable").table_style = "TableStyleMedium9"\n'
            'ws.range("B2:C13").number_format = "$#,##0"\n'
            "chart = ws.charts.add(400, 10, 500, 300)\n"
            'chart.set_source_data(ws.range("A1:C13"))\n'
            'chart.chart_type = "column_clustered"\n'
            'chart.name = "Sales vs Expenses"\n'
            'ws.range("A1:C1").column_width = 15\n'
        ),
    },
    "Employee Directory": {
        "description": "Employee directory with table, dropdowns for department, hyperlinks for email",
        "code": (
            'ws.range("A1:E1").value = ["Name", "Department", "Email", "Phone", "Start Date"]\n'
            'ws.range("A1:E1").font.bold = True\n'
            'ws.range("A1:E1").color = (44, 62, 80)\n'
            'ws.range("A1:E1").font.color = (255, 255, 255)\n'
            'ws.range("A2:E4").value = [["Alice Smith", "Engineering", "alice@example.com", "555-0101", "2022-01-15"], ["Bob Jones", "Marketing", "bob@example.com", "555-0102", "2021-06-01"], ["Carol Lee", "Engineering", "carol@example.com", "555-0103", "2023-03-20"]]\n'
            'ws.range("B2:B4").api.Validation.Add(Type=3, Formula1="Engineering,Marketing,Sales,HR,Finance")\n'
            'ws.range("B2:B4").api.Validation.InCellDropdown = True\n'
            'ws.range("C2:C4").add_hyperlink("mailto:alice@example.com", "alice@example.com")\n'
            'ws.tables.add(source=ws.range("A1:E4"), name="EmployeeTable").table_style = "TableStyleMedium2"\n'
            'ws.range("E2:E4").number_format = "YYYY-MM-DD"\n'
            'ws.freeze_panes.freeze_at("A2")\n'
            'ws.range("A1:E1").column_width = 20\n'
        ),
    },
    "Budget Tracker": {
        "description": "Monthly budget tracker with conditional formatting for over-budget items and totals row",
        "code": (
            'ws.range("A1:D1").value = ["Category", "Budget", "Actual", "Remaining"]\n'
            'ws.range("A1:D1").font.bold = True\n'
            'ws.range("A1:D1").color = (41, 128, 185)\n'
            'ws.range("A1:D1").font.color = (255, 255, 255)\n'
            'ws.range("A2:D6").value = [["Rent", 1500, 1500, "=B2-C2"], ["Food", 500, 620, "=B3-C3"], ["Transport", 200, 180, "=B4-C4"], ["Entertainment", 150, 230, "=B5-C5"], ["Utilities", 300, 280, "=B6-C6"]]\n'
            'ws.range("B2:D6").number_format = "$#,##0.00"\n'
            'ws.range("D2:D6").api.FormatConditions.Add(Type=1, Operator=6, Formula1="0")\n'
            'ws.range("D2:D6").api.FormatConditions(1).Font.Color = 255\n'
            'ws.range("D2:D6").api.FormatConditions(1).Interior.Color = 13551615\n'
            'ws.range("A7").value = "TOTAL"\n'
            'ws.range("A7").font.bold = True\n'
            'ws.range("B7").formula = "=SUM(B2:B6)"\n'
            'ws.range("C7").formula = "=SUM(C2:C6)"\n'
            'ws.range("D7").formula = "=SUM(D2:D6)"\n'
            'ws.range("B7:D7").number_format = "$#,##0.00"\n'
            "for i in range(1, 7):\n"
            '    ws.range("A1:D7").api.Borders(i).LineStyle = 1\n'
            'ws.range("A1:D1").column_width = 18\n'
        ),
    },
    "Product Catalog": {
        "description": "Product catalog with images placeholder, currency formatting, and hyperlinks",
        "code": (
            'ws.range("A1:E1").value = ["Product", "Category", "Price", "Stock", "Link"]\n'
            'ws.range("A1:E1").font.bold = True\n'
            'ws.range("A1:E1").color = (192, 57, 43)\n'
            'ws.range("A1:E1").font.color = (255, 255, 255)\n'
            'ws.range("A2:E5").value = [["Widget A", "Electronics", 29.99, 150, "https://example.com/widget-a"], ["Gadget B", "Tools", 49.99, 75, "https://example.com/gadget-b"], ["Device C", "Electronics", 199.99, 30, "https://example.com/device-c"], ["Tool D", "Tools", 15.99, 200, "https://example.com/tool-d"]]\n'
            'ws.range("C2:C5").number_format = "$#,##0.00"\n'
            "for i in range(2, 6):\n"
            '    ws.range(f"E{i}").add_hyperlink(ws.range(f"E{i}").value, "View Product")\n'
            'ws.tables.add(source=ws.range("A1:E5"), name="ProductTable").table_style = "TableStyleMedium4"\n'
            'ws.range("A1:E1").column_width = 18\n'
        ),
    },
    "Grade Sheet": {
        "description": "Student grade sheet with conditional formatting for pass/fail and color scale",
        "code": (
            'ws.range("A1:E1").value = ["Student", "Midterm", "Final", "Average", "Grade"]\n'
            'ws.range("A1:E1").font.bold = True\n'
            'ws.range("A1:E1").color = (39, 174, 96)\n'
            'ws.range("A1:E1").font.color = (255, 255, 255)\n'
            'ws.range("A2:E5").value = [["Alice", 85, 90, "=(B2+C2)/2", ""], ["Bob", 55, 60, "=(B3+C3)/2", ""], ["Carol", 92, 95, "=(B4+C4)/2", ""], ["Dave", 40, 50, "=(B5+C5)/2", ""]]\n'
            'ws.range("E2").formula = \'=IF(D2>=70,"Pass","Fail")\'\n'
            'ws.range("E3").formula = \'=IF(D3>=70,"Pass","Fail")\'\n'
            'ws.range("E4").formula = \'=IF(D4>=70,"Pass","Fail")\'\n'
            'ws.range("E5").formula = \'=IF(D5>=70,"Pass","Fail")\'\n'
            'ws.range("D2:D5").api.FormatConditions.AddColorScale(3)\n'
            'ws.range("E2:E5").api.FormatConditions.Add(Type=1, Operator=3, Formula1="\\"Pass\\"")\n'
            'ws.range("E2:E5").api.FormatConditions(2).Interior.Color = 65280\n'
            'ws.freeze_panes.freeze_at("A2")\n'
            'ws.range("A1:E1").column_width = 15\n'
        ),
    },
}

TEMPLATE_NAMES = list(TEMPLATES.keys())
