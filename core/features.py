def build_test_case_prompt(requirement, priority, include_negative, include_boundary):
    parts = [
        "Generate a comprehensive set of test cases for the following requirement.",
        f"Requirement: {requirement}",
        f"Priority: {priority}",
    ]
    parts.append("\nWrite the test cases into the active Excel sheet using xlwings.")
    parts.append("Create a table with columns: Test Case ID, Description, Steps, Expected Result, Priority, Type (Positive/Negative/Boundary).")
    parts.append("Use bold headers with a blue background (#89b4fa) and white text.")
    parts.append("Apply conditional formatting: highlight 'High' priority rows in red, 'Medium' in yellow, 'Low' in green.")
    if include_negative:
        parts.append("Include negative test cases that verify the system handles invalid inputs, errors, and edge failures correctly.")
    if include_boundary:
        parts.append("Include boundary test cases that test the minimum, maximum, and just beyond typical value ranges.")
    parts.append("Start the data at A1. Use ws for the active sheet and wb for the workbook.")
    parts.append("In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.")
    parts.append("Respond with ONLY Python code. No explanation, no markdown, no backticks.")
    return "\n".join(parts)


def build_bug_report_prompt(description, module, severity):
    parts = [
        "Create a structured bug report in Excel for the following issue.",
        f"Bug Description: {description}",
        f"Module: {module}",
        f"Severity: {severity}",
        "",
        "Create a bug report table with columns: Bug ID, Title, Description, Module, Severity, Status, Steps to Reproduce, Expected Behavior, Actual Behavior, Reporter, Date.",
        "Add a summary section above the table showing: Total Bugs, Open Bugs, Critical Bugs.",
        "Format the severity column with conditional formatting: Critical = red, High = orange, Medium = yellow, Low = green.",
        "Use bold headers with dark background. Apply borders to the entire table.",
        "Start the data at A1. Use ws for the active sheet and wb for the workbook.",
        "In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.",
        "Respond with ONLY Python code. No explanation, no markdown, no backticks.",
    ]
    return "\n".join(parts)


def build_reconciliation_prompt(bank_sheet, ledger_sheet, tolerance):
    parts = [
        "Create a bank reconciliation worksheet that matches bank statement entries with ledger entries.",
        f"Bank Statement Sheet: {bank_sheet}",
        f"Ledger Sheet: {ledger_sheet}",
        f"Tolerance for matching amounts: {tolerance}",
        "",
        "Create a Reconciliation sheet with columns: Date, Bank Amount, Ledger Amount, Difference, Status (Matched/Unmatched/Variance).",
        "Read data from both sheets, match by date and amount within tolerance.",
        "Highlight unmatched rows in red and variance rows in yellow.",
        "Add a summary at the top: Total Bank, Total Ledger, Difference, Matched Count, Unmatched Count.",
        "Format the summary with bold text and a colored background.",
        "Start the data at A1. Use ws for the active sheet and wb for the workbook.",
        "In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.",
        "Respond with ONLY Python code. No explanation, no markdown, no backticks.",
    ]
    return "\n".join(parts)


def build_gantt_prompt(tasks_text, start_date):
    parts = [
        "Create a Gantt chart in Excel for the following project tasks.",
        f"Tasks (one per line, format: Task Name | Duration Days | Dependency):",
        tasks_text,
        f"Project Start Date: {start_date}",
        "",
        "Create a Gantt chart with:",
        "- Column A: Task Name",
        "- Column B: Start Date",
        "- Column C: End Date",
        "- Column D: Duration (days)",
        "- Column E onwards: Timeline bars (one column per day/week)",
        "Fill timeline cells with blue (#89b4fa) where the task is active.",
        "Use conditional formatting or direct cell coloring for the Gantt bars.",
        "Bold headers. Freeze the first row and first column.",
        "Format dates as DD/MM/YYYY.",
        "Start the data at A1. Use ws for the active sheet and wb for the workbook.",
        "In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.",
        "Respond with ONLY Python code. No explanation, no markdown, no backticks.",
    ]
    return "\n".join(parts)


def build_gradebook_prompt(num_students, assignments, grading_scale, passing_grade):
    parts = [
        f"Create a gradebook for {num_students} students with {assignments} assignments.",
        f"Grading Scale: {grading_scale}",
        f"Passing Grade: {passing_grade}",
        "",
        "Create a table with columns: Student ID, Student Name, then one column per Assignment (Assignment 1, Assignment 2, ...), Total Score, Average, Letter Grade, Pass/Fail.",
        "Fill in sample data with realistic scores (0-100).",
        "Add formulas for Total (SUM), Average (AVERAGE), Letter Grade (IF nested or VLOOKUP), and Pass/Fail (IF).",
        "Apply conditional formatting: scores below passing grade in red, 90+ in green.",
        "Add a class summary row at the bottom with averages per assignment.",
        "Bold headers with blue background. Apply borders to the entire table.",
        "Start the data at A1. Use ws for the active sheet and wb for the workbook.",
        "In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.",
        "Respond with ONLY Python code. No explanation, no markdown, no backticks.",
    ]
    return "\n".join(parts)


def build_emi_prompt(principal, rate, tenure, include_amortization, compare_rates):
    parts = [
        f"Create an EMI (Equated Monthly Installment) calculator.",
        f"Principal Amount: {principal}",
        f"Annual Interest Rate: {rate}%",
        f"Loan Tenure: {tenure} months",
        "",
        "Create an EMI calculation sheet with: Principal, Rate, Tenure, EMI (using PMT formula or mathematical calculation).",
        "EMI Formula: EMI = P * r * (1+r)^n / ((1+r)^n - 1) where r = monthly rate, n = tenure months.",
    ]
    if include_amortization:
        parts.append(
            "Create an amortization schedule table with columns: Month, Opening Balance, EMI, Interest, Principal Repayment, Closing Balance."
        )
        parts.append("Fill all rows for the entire tenure. Format currency columns with $#,##0.00.")
    if compare_rates:
        parts.append(
            "Add a comparison table showing EMI for different interest rates (current rate +/- 1%, 2%) side by side."
        )
    parts.append("Bold headers. Format amounts as currency. Apply borders.")
    parts.append("Start the data at A1. Use ws for the active sheet and wb for the workbook.")
    parts.append("In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.")
    parts.append("Respond with ONLY Python code. No explanation, no markdown, no backticks.")
    return "\n".join(parts)


def build_payroll_prompt(basic_col, hra_pct, da_pct, pf_pct, esi_pct, tds_pct, include_overtime):
    parts = [
        "Create a payroll calculation sheet for employees.",
        f"Basic Salary Column: {basic_col}",
        f"HRA: {hra_pct}% of Basic",
        f"DA: {da_pct}% of Basic",
        f"PF: {pf_pct}% of Basic",
        f"ESI: {esi_pct}% of Gross",
        f"TDS: {tds_pct}% of Taxable Income",
        "",
        "Create columns: Employee ID, Name, Basic, HRA, DA, Gross Salary, PF, ESI, TDS, Total Deductions, Net Salary.",
        "Add formulas for all calculated columns.",
    ]
    if include_overtime:
        parts.append("Add columns: Overtime Hours, Overtime Rate, Overtime Pay. Add OT Pay to Gross Salary.")
    parts.append("Add a summary row at the bottom with totals and averages.")
    parts.append("Format salary columns as currency ($#,##0.00). Bold headers with blue background.")
    parts.append("Conditional formatting: Net Salary below minimum wage in red.")
    parts.append("Start the data at A1. Use ws for the active sheet and wb for the workbook.")
    parts.append("In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.")
    parts.append("Respond with ONLY Python code. No explanation, no markdown, no backticks.")
    return "\n".join(parts)


def build_inventory_dashboard_prompt(product_col, stock_col, reorder_col, price_col, include_abc, low_stock_alerts):
    parts = [
        "Create an inventory management dashboard in Excel.",
        f"Product Name Column reference: {product_col}",
        f"Current Stock Column reference: {stock_col}",
        f"Reorder Level Column reference: {reorder_col}",
        f"Unit Price Column reference: {price_col}",
        "",
        "Create a table with: Product, Current Stock, Reorder Level, Unit Price, Stock Value, Status.",
        "Stock Value = Stock * Price. Status = IF(Current < Reorder, 'REORDER', 'OK').",
    ]
    if include_abc:
        parts.append(
            "Add ABC Analysis column: classify products as A (top 80% value), B (next 15%), C (bottom 5%) "
            "based on cumulative stock value. Use VLOOKUP or IF logic."
        )
    if low_stock_alerts:
        parts.append(
            "Apply conditional formatting: highlight 'REORDER' status rows in red, "
            "'OK' rows in green. Also highlight items where stock is less than 20% of reorder level in bold red."
        )
    parts.append("Add a dashboard summary at top: Total Products, Total Value, Items to Reorder, Out of Stock.")
    parts.append("Create a bar chart showing stock levels per product.")
    parts.append("Bold headers. Apply borders. Freeze top row.")
    parts.append("Start the data at A1. Use ws for the active sheet and wb for the workbook.")
    parts.append("In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.")
    parts.append("Respond with ONLY Python code. No explanation, no markdown, no backticks.")
    return "\n".join(parts)


def build_data_cleaning_prompt(sheet_name, options: dict):
    parts = [
        f"Clean the data in the sheet '{sheet_name}' using the following operations:",
        "",
    ]
    if options.get("remove_duplicates"):
        parts.append(
            "1. REMOVE DUPLICATES: Read all data, identify duplicate rows based on all columns, "
            "and remove them. Keep the first occurrence."
        )
    if options.get("trim_whitespace"):
        parts.append(
            "2. TRIM WHITESPACE: Read all text cells, strip leading/trailing whitespace, "
            "and collapse multiple internal spaces to single spaces."
        )
    if options.get("fix_dates"):
        parts.append(
            "3. FIX DATE FORMATS: Identify date columns and standardize to YYYY-MM-DD format. "
            "Detect DD/MM/YYYY vs MM/DD/YYYY ambiguity using context clues (values > 12 in the day position)."
        )
    if options.get("standardize_phones"):
        parts.append(
            "4. STANDARDIZE PHONE NUMBERS: Normalize phone numbers to a consistent format "
            "(e.g., +1-XXX-XXX-XXXX for US, +91-XXXXX-XXXXX for India). "
            "Remove parentheses, dashes, and spaces before reformatting."
        )
    if options.get("normalize_text"):
        parts.append(
            "5. NORMALIZE TEXT: Apply Title Case to name columns. Apply UPPER CASE to category/status columns. "
            "Trim and normalize all text fields."
        )
    if options.get("remove_empty_rows"):
        parts.append(
            "6. REMOVE EMPTY ROWS: Scan all rows and delete any row where all cells are empty or None."
        )
    if options.get("fill_down"):
        parts.append(
            "7. FILL DOWN MISSING VALUES: For each column, fill None/empty cells with the last non-empty "
            "value above (forward fill). Skip the header row."
        )
    parts.append("")
    parts.append("Read the current sheet data, apply all selected operations in order, and write the cleaned data back.")
    parts.append("Add a summary row at the bottom showing: Original Rows, Cleaned Rows, Duplicates Removed, Empty Rows Removed.")
    parts.append("Bold headers. Apply borders.")
    parts.append("Start the data at A1. Use ws for the active sheet and wb for the workbook.")
    parts.append("In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.")
    parts.append("Respond with ONLY Python code. No explanation, no markdown, no backticks.")
    return "\n".join(parts)


def build_health_check_prompt(sheet_data):
    import json

    preview = sheet_data[:25] if sheet_data else []
    data_str = json.dumps(preview, indent=2, ensure_ascii=False) if preview else "(empty sheet)"

    parts = [
        "Audit this spreadsheet data and identify the following issues:",
        "",
        "1. BROKEN REFERENCES: Cells with formulas referencing non-existent sheets, ranges, or #REF! errors.",
        "2. HARDCODED VALUES: Numbers that should be formulas (e.g., totals that are typed instead of SUM).",
        "3. INCONSISTENT FORMATTING: Mixed date formats, number formats, or text casing in the same column.",
        "4. MISSING DATA: Cells that are empty but should have values based on surrounding context.",
        "5. POTENTIAL ERRORS: Values that are statistical outliers (e.g., negative ages, impossible dates).",
        "6. CIRCULAR REFERENCES: Formulas that reference their own cell directly or indirectly.",
        "",
        f"Sheet data (first {len(preview)} rows):",
        data_str,
        "",
        "Write Python xlwings analysis code that reads the sheet data, checks for each issue category, "
        "and prints a detailed report. Use print() for all output.",
        "The sheet object is 'ws' and workbook is 'wb'.",
        "Respond with ONLY Python code. No explanation, no markdown, no backticks.",
    ]
    return "\n".join(parts)


def build_vba_to_python_prompt(vba_code):
    parts = [
        "Convert this VBA macro to Python xlwings code. The sheet object is 'ws' and workbook is 'wb'.",
        "",
        "VBA Code:",
        vba_code,
        "",
        "Conversion rules:",
        "- Replace Range(\"A1\") with ws.range(\"A1\")",
        "- Replace ActiveSheet with ws",
        "- Replace ActiveWorkbook with wb",
        "- Replace Cells(row, col) with ws.range((row, col))",
        "- Replace Worksheets(\"Name\") with wb.sheets['Name']",
        "- Replace For/Next with Python for loops",
        "- Replace If/Then/Else with Python if/elif/else",
        "- Replace Dim with Python variable assignments (no declaration needed)",
        "- Replace MsgBox with print()",
        "- Use ws.range().value for reading and ws.range().value = for writing",
        "- Use ws.range().formula for formula operations",
        "- Use ws.range().color = (R, G, B) for cell colors",
        "- Use ws.range().font.bold, .font.size, etc. for formatting",
        "",
        "Respond with ONLY Python code. No explanation, no markdown, no backticks.",
    ]
    return "\n".join(parts)


def build_sentiment_prompt(column_name, output_format, max_rows):
    parts = [
        "Perform sentiment analysis on the text in the following column of the current Excel sheet.",
        f"Column Name: {column_name}",
        f"Output Format: {output_format}",
        f"Maximum Rows to Process: {max_rows}",
        "",
        "Read the data from the specified column, analyze the sentiment of each text entry, "
        "and write the results in a new column called 'Sentiment'.",
        "Sentiment values should be: Positive, Negative, or Neutral.",
        "If output format includes 'Score', also add a 'Sentiment Score' column with values from -1.0 to 1.0.",
        "Apply conditional formatting: Positive = green, Negative = red, Neutral = yellow.",
        "Add a summary row at the bottom: Total Positive, Total Negative, Total Neutral counts.",
        "Bold headers. Apply borders.",
        "Start the data at A1. Use ws for the active sheet and wb for the workbook.",
        "In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.",
        "Respond with ONLY Python code. No explanation, no markdown, no backticks.",
    ]
    return "\n".join(parts)


def build_raci_prompt(phases, team_members):
    parts = [
        "Create a RACI matrix (Responsible, Accountable, Consulted, Informed) for the following project.",
        f"Project Phases: {phases}",
        f"Team Members: {team_members}",
        "",
        "Create a matrix with phases as rows and team members as columns.",
        "Fill in R (Responsible), A (Accountable), C (Consulted), I (Informed) for each cell.",
        "Each row must have exactly one 'A' (Accountable) and at least one 'R' (Responsible).",
        "Apply conditional formatting: R = blue, A = red, C = yellow, I = green.",
        "Center-align all RACI values. Bold headers and phase names.",
        "Add borders to the entire matrix. Freeze the first column.",
        "Start the data at A1. Use ws for the active sheet and wb for the workbook.",
        "In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.",
        "Respond with ONLY Python code. No explanation, no markdown, no backticks.",
    ]
    return "\n".join(parts)


def build_onboarding_prompt(role, department, duration):
    parts = [
        "Create an employee onboarding checklist for the following role.",
        f"Role: {role}",
        f"Department: {department}",
        f"Duration: {duration}",
        "",
        "Create a table with columns: Day/Week, Task, Category (IT Setup/HR/Paperwork/Training/Team Intro), "
        "Responsible Person, Status (Pending/In Progress/Complete), Notes.",
        "Organize tasks chronologically by day/week for the full duration.",
        "Apply conditional formatting: Pending = yellow, In Progress = blue, Complete = green.",
        "Add a progress summary section at the top showing completion percentage.",
        "Bold headers with colored background. Apply borders. Freeze top row.",
        "Start the data at A1. Use ws for the active sheet and wb for the workbook.",
        "In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.",
        "Respond with ONLY Python code. No explanation, no markdown, no backticks.",
    ]
    return "\n".join(parts)


def build_flashcard_prompt(terms_text, randomize, quiz_mode):
    parts = [
        "Create flashcards in Excel for the following terms and definitions.",
        f"Terms: {terms_text}",
        "",
        "Create a table with columns: Card #, Term, Definition, Category, Difficulty (1-5).",
    ]
    if randomize:
        parts.append("Randomize the order of the flashcards using random.shuffle().")
    if quiz_mode:
        parts.append(
            "Add a 'Quiz Answer' column where the definition is hidden and a 'Revealed' column "
            "where the definition is shown. Mark Quiz Answer as '???'."
        )
    parts.append("Apply alternating row colors for readability. Bold the Term column.")
    parts.append("Start the data at A1. Use ws for the active sheet and wb for the workbook.")
    parts.append("In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.")
    parts.append("Respond with ONLY Python code. No explanation, no markdown, no backticks.")
    return "\n".join(parts)


def build_study_planner_prompt(subjects, exam_date, study_days, style):
    parts = [
        "Create a study plan in Excel for the following subjects and exam schedule.",
        f"Subjects: {subjects}",
        f"Exam Date: {exam_date}",
        f"Days Until Exam: {study_days}",
        f"Study Style: {style} (Pomodoro/Spaced Repetition/Block Schedule)",
        "",
        "Create a calendar-style study plan with:",
        "- Column A: Date",
        "- Column B: Day of Week",
        "- Column C: Subject",
        "- Column D: Topic/Chapter",
        "- Column E: Duration (hours)",
        "- Column F: Study Method",
        "- Column G: Completion Status",
        "Distribute study time based on subject difficulty and days until exam.",
        "Include rest days and review sessions.",
        "Apply conditional formatting: Completed = green, In Progress = blue, Not Started = red.",
        "Bold headers. Apply borders. Freeze top row.",
        "Start the data at A1. Use ws for the active sheet and wb for the workbook.",
        "In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.",
        "Respond with ONLY Python code. No explanation, no markdown, no backticks.",
    ]
    return "\n".join(parts)


def build_risk_register_prompt(project_name, num_categories):
    parts = [
        f"Create a risk register for the project: {project_name}",
        f"Number of risk categories: {num_categories}",
        "",
        "Create a table with columns: Risk ID, Category, Risk Description, Likelihood (1-5), "
        "Impact (1-5), Risk Score (Likelihood * Impact), Mitigation Strategy, Owner, Status.",
        "Generate realistic risks across the specified number of categories.",
        "Apply conditional formatting: Risk Score >= 15 = red, 8-14 = yellow, < 8 = green.",
        "Add a risk heat map summary showing likelihood vs impact matrix.",
        "Bold headers. Apply borders. Freeze top row.",
        "Start the data at A1. Use ws for the active sheet and wb for the workbook.",
        "In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.",
        "Respond with ONLY Python code. No explanation, no markdown, no backticks.",
    ]
    return "\n".join(parts)


def build_sprint_backlog_prompt(sprint_name, duration_weeks, velocity, stories):
    parts = [
        f"Create a sprint backlog for: {sprint_name}",
        f"Sprint Duration: {duration_weeks} weeks",
        f"Team Velocity: {velocity} story points",
        f"User Stories: {stories}",
        "",
        "Create a table with columns: Story ID, User Story, Story Points, Priority, "
        "Status (To Do/In Progress/In Review/Done), Assignee, Tasks Breakdown.",
        "Distribute stories within the sprint velocity. Add realistic task breakdowns.",
        "Apply conditional formatting: Done = green, In Progress = blue, In Review = yellow, To Do = red.",
        "Add a sprint summary: Total Stories, Total Points, Points Completed, Burndown projection.",
        "Bold headers. Apply borders. Freeze top row.",
        "Start the data at A1. Use ws for the active sheet and wb for the workbook.",
        "In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.",
        "Respond with ONLY Python code. No explanation, no markdown, no backticks.",
    ]
    return "\n".join(parts)


def build_lead_scoring_prompt(company_size_col, engagement_col, industry_col, budget_col, rules):
    parts = [
        "Create a lead scoring model in Excel.",
        f"Company Size Column: {company_size_col}",
        f"Engagement Level Column: {engagement_col}",
        f"Industry Column: {industry_col}",
        f"Budget Column: {budget_col}",
        f"Scoring Rules: {rules}",
        "",
        "Create a table with: Lead Name, Company Size, Engagement, Industry, Budget, "
        "Size Score, Engagement Score, Industry Score, Budget Score, Total Score, Grade (A/B/C/D).",
        "Add scoring columns that assign points based on the rules provided.",
        "Total Score = sum of all individual scores. Grade based on total score thresholds.",
        "Apply conditional formatting: A = green, B = blue, C = yellow, D = red.",
        "Sort by Total Score descending. Add a summary of lead distribution by grade.",
        "Bold headers. Apply borders. Freeze top row.",
        "Start the data at A1. Use ws for the active sheet and wb for the workbook.",
        "In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.",
        "Respond with ONLY Python code. No explanation, no markdown, no backticks.",
    ]
    return "\n".join(parts)


def build_campaign_roi_prompt(name_col, spend_col, clicks_col, conversions_col, revenue_col):
    parts = [
        "Create a marketing campaign ROI calculation sheet.",
        f"Campaign Name Column: {name_col}",
        f"Spend Column: {spend_col}",
        f"Clicks Column: {clicks_col}",
        f"Conversions Column: {conversions_col}",
        f"Revenue Column: {revenue_col}",
        "",
        "Create calculated columns: CPC (Spend/Clicks), Conversion Rate (Conversions/Clicks), "
        "CPA (Spend/Conversions), ROAS (Revenue/Spend), ROI ((Revenue-Spend)/Spend*100), Profit (Revenue-Spend).",
        "Add conditional formatting: ROI > 100% = green, 0-100% = yellow, < 0% = red.",
        "Format currency columns as $#,##0.00. Format percentages as 0.00%.",
        "Add a summary row with totals and averages. Create a bar chart comparing Spend vs Revenue.",
        "Bold headers. Apply borders. Freeze top row.",
        "Start the data at A1. Use ws for the active sheet and wb for the workbook.",
        "In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.",
        "Respond with ONLY Python code. No explanation, no markdown, no backticks.",
    ]
    return "\n".join(parts)


def build_email_cleaner_prompt(options: dict):
    parts = [
        "Clean an email list in the active Excel sheet with the following operations:",
    ]
    if options.get("remove_duplicates"):
        parts.append("- Remove duplicate email addresses")
    if options.get("fix_case"):
        parts.append("- Convert all emails to lowercase")
    if options.get("remove_invalid"):
        parts.append("- Remove emails that don't match standard email format (contains @ and domain)")
    if options.get("trim_whitespace"):
        parts.append("- Trim leading/trailing whitespace from email addresses")
    if options.get("remove_unsubscribed"):
        parts.append("- Mark or remove unsubscribed/bounced emails (mark in a Status column)")
    if options.get("categorize_domain"):
        parts.append("- Add a Domain column extracting the domain from each email")
    parts.append("")
    parts.append("Read the email list, apply all operations, and write back the cleaned data.")
    parts.append("Add a summary: Original Count, After Cleaning, Duplicates Removed, Invalid Removed.")
    parts.append("Bold headers. Apply borders.")
    parts.append("Start the data at A1. Use ws for the active sheet and wb for the workbook.")
    parts.append("In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.")
    parts.append("Respond with ONLY Python code. No explanation, no markdown, no backticks.")
    return "\n".join(parts)


def build_reorder_point_prompt(product_col, demand_col, lead_time_col, service_level, safety_factor):
    parts = [
        "Calculate reorder points for inventory items.",
        f"Product Column: {product_col}",
        f"Average Daily Demand Column: {demand_col}",
        f"Lead Time (days) Column: {lead_time_col}",
        f"Service Level: {service_level}%",
        f"Safety Factor (Z-score): {safety_factor}",
        "",
        "Create calculated columns:",
        "- Average Demand During Lead Time = Demand * Lead Time",
        "- Safety Stock = Safety Factor * SQRT(Lead Time) * Demand Std Dev (estimate if not available)",
        "- Reorder Point = Demand During Lead Time + Safety Stock",
        "- Status = IF(Current Stock <= Reorder Point, 'REORDER NOW', 'OK')",
        "Apply conditional formatting: REORDER NOW = red, OK = green.",
        "Format number columns with appropriate precision. Bold headers. Apply borders.",
        "Start the data at A1. Use ws for the active sheet and wb for the workbook.",
        "In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.",
        "Respond with ONLY Python code. No explanation, no markdown, no backticks.",
    ]
    return "\n".join(parts)


def build_rental_yield_prompt(price_col, rent_col, maintenance_col, vacancy_rate):
    parts = [
        "Calculate rental yield and investment metrics for properties.",
        f"Property Price Column: {price_col}",
        f"Monthly Rent Column: {rent_col}",
        f"Annual Maintenance Column: {maintenance_col}",
        f"Vacancy Rate: {vacancy_rate}%",
        "",
        "Create calculated columns:",
        "- Annual Rent = Monthly Rent * 12",
        "- Effective Annual Rent = Annual Rent * (1 - Vacancy Rate/100)",
        "- Net Operating Income = Effective Annual Rent - Maintenance",
        "- Gross Yield = (Annual Rent / Price) * 100",
        "- Net Yield = (NOI / Price) * 100",
        "- Cap Rate = Net Yield (same calculation, different convention)",
        "Apply conditional formatting: Net Yield > 5% = green, 3-5% = yellow, < 3% = red.",
        "Format yield columns as percentages. Format currency as $#,##0.00.",
        "Add a summary row with averages. Bold headers. Apply borders.",
        "Start the data at A1. Use ws for the active sheet and wb for the workbook.",
        "In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.",
        "Respond with ONLY Python code. No explanation, no markdown, no backticks.",
    ]
    return "\n".join(parts)


def build_clinical_cleaner_prompt(options: dict):
    parts = [
        "Clean clinical/medical data in the active Excel sheet with the following operations:",
    ]
    if options.get("standardize_codes"):
        parts.append("- Standardize ICD-10 and CPT codes to proper format (e.g., A00.0 for ICD, 00000 for CPT)")
    if options.get("fix_dates"):
        parts.append("- Standardize all date fields to ISO 8601 format (YYYY-MM-DD)")
    if options.get("normalize_units"):
        parts.append("- Normalize measurement units (e.g., all weights in kg, heights in cm, temperatures in Celsius)")
    if options.get("remove_phi"):
        parts.append("- Remove or mask Protected Health Information (PHI): SSN, full names, addresses, phone numbers")
    if options.get("validate_ranges"):
        parts.append("- Flag values outside clinical normal ranges (e.g., heart rate 0-300, temperature 30-45°C, BP 50-300)")
    if options.get("fill_missing"):
        parts.append("- Fill missing values with 'N/A' for text and None for numeric fields")
    parts.append("")
    parts.append("Read the clinical data, apply all operations, and write back the cleaned data.")
    parts.append("Add a validation summary row showing: Total Records, Records with Issues, Fields Cleaned.")
    parts.append("Bold headers. Apply borders. Red font for flagged values.")
    parts.append("Start the data at A1. Use ws for the active sheet and wb for the workbook.")
    parts.append("In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.")
    parts.append("Respond with ONLY Python code. No explanation, no markdown, no backticks.")
    return "\n".join(parts)


def build_medication_tracker_prompt(num_meds, duration, side_effects, refill, time_of_day):
    parts = [
        f"Create a medication tracking sheet for {num_meds} medications over {duration}.",
        f"Track Side Effects: {'Yes' if side_effects else 'No'}",
        f"Include Refill Reminders: {'Yes' if refill else 'No'}",
        f"Time of Day: {time_of_day}",
        "",
        "Create a table with columns: Date, Medication Name, Dosage, Time Taken, "
        "Status (Taken/Missed/Skipped), Side Effects (if enabled: None/Mild/Moderate/Severe).",
    ]
    if refill:
        parts.append("Add a Refill Tracker section: Medication, Current Supply, Daily Dose, Days Remaining, Refill Date.")
        parts.append("Highlight medications with < 7 days supply in yellow, < 3 days in red.")
    parts.append("Apply conditional formatting: Taken = green, Missed = red, Skipped = yellow.")
    parts.append("Add a compliance summary: Total Scheduled, Total Taken, Compliance %.")
    parts.append("Bold headers. Apply borders. Freeze top row and first column.")
    parts.append("Start the data at A1. Use ws for the active sheet and wb for the workbook.")
    parts.append("In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.")
    parts.append("Respond with ONLY Python code. No explanation, no markdown, no backticks.")
    return "\n".join(parts)


def build_patient_schedule_prompt(start, end, duration, lunch_start, lunch_duration, max_patients, buffer):
    parts = [
        "Create a patient appointment schedule.",
        f"Clinic Hours: {start} to {end}",
        f"Appointment Duration: {duration} minutes",
        f"Lunch Break: {lunch_start} for {lunch_duration} minutes",
        f"Maximum Patients: {max_patients}",
        f"Buffer Between Appointments: {buffer} minutes",
        "",
        "Create a schedule table with columns: Slot #, Time, Patient Name, Type (New/Follow-up/Urgent), "
        "Duration, Status (Available/Booked/Cancelled), Notes.",
        "Generate all time slots for the day respecting the lunch break and buffer times.",
        "Apply conditional formatting: Available = green, Booked = blue, Cancelled = red.",
        "Bold headers. Apply borders. Center-align time columns.",
        "Add a summary: Total Slots, Available, Booked, Utilization %.",
        "Start the data at A1. Use ws for the active sheet and wb for the workbook.",
        "In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.",
        "Respond with ONLY Python code. No explanation, no markdown, no backticks.",
    ]
    return "\n".join(parts)


def build_tax_classifier_prompt(expense_col, include_gst, include_us):
    parts = [
        f"Classify expenses in column '{expense_col}' for tax purposes.",
        f"Include GST Categories: {'Yes' if include_gst else 'No'}",
        f"Include US Tax Categories: {'Yes' if include_us else 'No'}",
        "",
        "Create columns: Expense Description, Category (Travel/Meals/Office/Software/Marketing/etc.), "
        "Tax Dedible (Yes/No/Partial), Deduction %, Tax Category Code.",
    ]
    if include_gst:
        parts.append(
            "Add GST classification: GST Applicable (Yes/No), GST Rate (5%/12%/18%/28%), "
            "Input Tax Credit Available (Yes/No)."
        )
    if include_us:
        parts.append(
            "Add US tax categories: Schedule C category, IRS Category Code, "
            "Ordinary & Necessary (Yes/No)."
        )
    parts.append("Apply conditional formatting: Fully Deductible = green, Partial = yellow, Not Deductible = red.")
    parts.append("Bold headers. Apply borders.")
    parts.append("Start the data at A1. Use ws for the active sheet and wb for the workbook.")
    parts.append("In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.")
    parts.append("Respond with ONLY Python code. No explanation, no markdown, no backticks.")
    return "\n".join(parts)


def build_traceability_prompt(requirements, test_case_ids):
    parts = [
        "Create a requirements traceability matrix.",
        f"Requirements: {requirements}",
        f"Test Case IDs: {test_case_ids}",
        "",
        "Create a matrix with requirements as rows and test cases as columns.",
        "Mark 'X' where a test case covers a requirement.",
        "Add coverage columns: # Test Cases Linked, Coverage Status (Full/Partial/None).",
        "Apply conditional formatting: Full = green, Partial = yellow, None = red.",
        "Add a coverage summary row at the bottom.",
        "Bold headers. Center-align X marks. Apply borders. Freeze first column.",
        "Start the data at A1. Use ws for the active sheet and wb for the workbook.",
        "In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.",
        "Respond with ONLY Python code. No explanation, no markdown, no backticks.",
    ]
    return "\n".join(parts)


def build_test_data_prompt(data_type, count, include_invalid, international):
    parts = [
        f"Generate {count} rows of realistic test data for: {data_type}.",
        f"Include Invalid/Edge Cases: {'Yes' if include_invalid else 'No'}",
        f"International Formats: {'Yes' if international else 'No'}",
        "",
        "Create appropriate columns based on the data type (e.g., for 'users': ID, Name, Email, Phone, Address, DOB).",
    ]
    if include_invalid:
        parts.append(
            "Include ~10% invalid data rows: malformed emails, impossible dates, "
            "negative numbers for positive-only fields, SQL injection strings, XSS payloads."
        )
    if international:
        parts.append(
            "Include international data: non-ASCII names (e.g., José, Müller, 李), "
            "international phone formats, various date formats, multiple currencies."
        )
    parts.append("Bold headers. Apply borders. Format dates as YYYY-MM-DD.")
    parts.append("Start the data at A1. Use ws for the active sheet and wb for the workbook.")
    parts.append("In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.")
    parts.append("Respond with ONLY Python code. No explanation, no markdown, no backticks.")
    return "\n".join(parts)


def build_property_comparison_prompt(num_properties, criteria, include_roi):
    parts = [
        f"Create a property comparison sheet for {num_properties} properties.",
        f"Comparison Criteria: {criteria}",
        f"Include ROI Calculation: {'Yes' if include_roi else 'No'}",
        "",
        "Create columns: Property Name, Location, Price, Size (sq ft), Bedrooms, Bathrooms, "
        "Price/sqft, Year Built, Parking, Score.",
    ]
    if include_roi:
        parts.append(
            "Add ROI columns: Estimated Monthly Rent, Annual Rent, Gross Yield, "
            "Property Tax, Insurance, Net Yield, 5-Year Appreciation Estimate."
        )
    parts.append("Calculate a composite Score (1-10) based on weighted criteria.")
    parts.append("Apply conditional formatting: Score >= 8 = green, 5-7 = yellow, < 5 = red.")
    parts.append("Add a summary comparison row. Bold headers. Apply borders. Freeze first column.")
    parts.append("Start the data at A1. Use ws for the active sheet and wb for the workbook.")
    parts.append("In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.")
    parts.append("Respond with ONLY Python code. No explanation, no markdown, no backticks.")
    return "\n".join(parts)


def build_invoice_extract_prompt(format_type):
    parts = [
        f"Create an invoice data extraction template for: {format_type}",
        "",
        "Create a structured invoice template with sections:",
        "- Invoice Header: Invoice #, Date, Due Date, Terms",
        "- From (Seller): Company, Address, Contact, Tax ID",
        "- To (Buyer): Company, Address, Contact, Tax ID",
        "- Line Items: #, Description, Quantity, Unit Price, Amount",
        "- Subtotal, Tax, Discount, Total",
        "- Payment Details: Bank, Account, Routing",
    ]
    if format_type == "Receipt":
        parts.append("Simplify for receipt format: Store, Date, Items, Subtotal, Tax, Total, Payment Method.")
    elif format_type == "Proforma":
        parts.append("Mark all amounts as estimated. Add validity period and terms.")
    parts.append("Apply number formatting for currency. Bold section headers. Apply borders.")
    parts.append("Start the data at A1. Use ws for the active sheet and wb for the workbook.")
    parts.append("In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.")
    parts.append("Respond with ONLY Python code. No explanation, no markdown, no backticks.")
    return "\n".join(parts)


def build_attendance_report_prompt(name_col, date_col, status_col, work_start, work_end, late_threshold):
    parts = [
        "Create an attendance analysis report from the current sheet data.",
        f"Name Column: {name_col}",
        f"Date Column: {date_col}",
        f"Status Column: {status_col}",
        f"Work Start Time: {work_start}",
        f"Work End Time: {work_end}",
        f"Late Threshold (minutes): {late_threshold}",
        "",
        "Read the attendance data and create a summary sheet with:",
        "- Employee Name",
        "- Total Working Days",
        "- Days Present",
        "- Days Absent",
        "- Days Late",
        "- Attendance % (Present/Total * 100)",
        "- Average Arrival Time",
        "- Status (Good >= 95%, Average 80-95%, Poor < 80%)",
        "Apply conditional formatting: Good = green, Average = yellow, Poor = red.",
        "Format attendance % as 0.0%. Bold headers. Apply borders.",
        "Create a pivot-style summary table.",
        "Start the data at A1. Use ws for the active sheet and wb for the workbook.",
        "In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.",
        "Respond with ONLY Python code. No explanation, no markdown, no backticks.",
    ]
    return "\n".join(parts)


def build_shipping_tracker_prompt(num_shipments, carrier_dropdown, auto_eta, status_format):
    parts = [
        f"Create a shipping tracker for {num_shipments} shipments.",
        f"Carrier Dropdown: {'Yes' if carrier_dropdown else 'No'}",
        f"Auto-Calculate ETA: {'Yes' if auto_eta else 'No'}",
        f"Status Format: {status_format}",
        "",
        "Create columns: Tracking #, Carrier, Origin, Destination, Ship Date, "
        "Estimated Delivery, Actual Delivery, Days in Transit, Status, Notes.",
        "Generate sample shipment data with realistic tracking numbers and routes.",
    ]
    if carrier_dropdown:
        parts.append(
            "Add data validation dropdown for Carrier column with: FedEx, UPS, DHL, USPS, BlueDart, DTDC."
        )
    if auto_eta:
        parts.append(
            "Calculate Estimated Delivery based on carrier average transit times: "
            "FedEx=3 days, UPS=4 days, DHL=5 days, USPS=7 days."
        )
    parts.append(f"Use status format: {status_format} (e.g., Shipped/In Transit/Delivered/Delayed/Returned).")
    parts.append("Apply conditional formatting: Delivered = green, In Transit = blue, Delayed = red.")
    parts.append("Bold headers. Apply borders. Format dates as DD/MM/YYYY.")
    parts.append("Start the data at A1. Use ws for the active sheet and wb for the workbook.")
    parts.append("In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.")
    parts.append("Respond with ONLY Python code. No explanation, no markdown, no backticks.")
    return "\n".join(parts)


def build_pl_variance_prompt(budget_sheet, actual_sheet, threshold):
    parts = [
        "Create a P&L Variance Analysis report comparing budget vs actual.",
        f"Budget Sheet: {budget_sheet}",
        f"Actual Sheet: {actual_sheet}",
        f"Variance Alert Threshold: {threshold}%",
        "",
        "Create a report with columns: Category, Line Item, Budget Amount, Actual Amount, "
        "Variance ($), Variance (%), Status (Under Budget/On Budget/Over Budget).",
        "Variance % = (Actual - Budget) / Budget * 100.",
        f"Flag items where absolute variance % exceeds {threshold}%.",
        "Apply conditional formatting: Over Budget = red, Under Budget = green, On Budget = blue.",
        "Format currency as $#,##0.00. Format percentages as 0.0%.",
        "Add subtotals for each category. Add grand total row.",
        "Create a bar chart comparing Budget vs Actual by category.",
        "Bold headers and category rows. Apply borders.",
        "Start the data at A1. Use ws for the active sheet and wb for the workbook.",
        "In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.",
        "Respond with ONLY Python code. No explanation, no markdown, no backticks.",
    ]
    return "\n".join(parts)


def build_directory_prompt(departments, photo_placeholder, filter_dropdown):
    parts = [
        f"Create an employee directory organized by departments: {departments}.",
        f"Include Photo Placeholder: {'Yes' if photo_placeholder else 'No'}",
        f"Add Filter Dropdown: {'Yes' if filter_dropdown else 'No'}",
        "",
        "Create columns: Employee ID, Full Name, Department, Job Title, Email, "
        "Phone, Office Location, Start Date.",
    ]
    if photo_placeholder:
        parts.append("Add a Photo column with placeholder text '[Photo]' centered in each cell.")
    if filter_dropdown:
        parts.append(
            "Add data validation dropdown filters for Department column using the specified department list."
        )
    parts.append("Generate sample employee data across all departments.")
    parts.append("Sort by Department then Name. Apply alternating row colors. Bold headers.")
    parts.append("Apply borders. Freeze top row. Center-align ID and Phone columns.")
    parts.append("Start the data at A1. Use ws for the active sheet and wb for the workbook.")
    parts.append("In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.")
    parts.append("Respond with ONLY Python code. No explanation, no markdown, no backticks.")
    return "\n".join(parts)


def build_status_report_prompt(task_sheet, project_name, period, include_risks):
    parts = [
        "Create a project status report.",
        f"Task Data Sheet: {task_sheet}",
        f"Project Name: {project_name}",
        f"Reporting Period: {period}",
        f"Include Risk Section: {'Yes' if include_risks else 'No'}",
        "",
        "Create a Status Report sheet with sections:",
        "1. Project Header: Project Name, Report Date, Period, Project Manager, Overall Status.",
        "2. Summary: Total Tasks, Completed, In Progress, Not Started, Completion %.",
        "3. Task Detail Table: Task ID, Task Name, Assignee, Status, % Complete, Due Date, Notes.",
        "4. Milestone Timeline: Milestone, Target Date, Actual Date, Status.",
    ]
    if include_risks:
        parts.append(
            "5. Risk Section: Risk ID, Description, Impact, Likelihood, Mitigation, Owner."
        )
    parts.append("Apply conditional formatting: On Track = green, At Risk = yellow, Behind = red.")
    parts.append("Format dates. Bold section headers with colored backgrounds. Apply borders.")
    parts.append("Start the data at A1. Use ws for the active sheet and wb for the workbook.")
    parts.append("In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.")
    parts.append("Respond with ONLY Python code. No explanation, no markdown, no backticks.")
    return "\n".join(parts)


def build_consolidator_prompt(file_count, remove_dupes, add_source, align_columns):
    parts = [
        f"Create a multi-file consolidation script for {file_count} Excel files.",
        f"Remove Duplicates: {'Yes' if remove_dupes else 'No'}",
        f"Add Source File Column: {'Yes' if add_source else 'No'}",
        f"Align Columns: {'Yes' if align_columns else 'No'}",
        "",
        "Write Python xlwings code to:",
        "1. Open each file and read all data from the first sheet.",
    ]
    if align_columns:
        parts.append(
            "2. Align columns by header name — if files have different column orders, "
            "reorder to match the first file's column structure."
        )
    else:
        parts.append("2. Concatenate all data assuming identical column structures.")
    if add_source:
        parts.append("3. Add a 'Source File' column to each row indicating which file it came from.")
    if remove_dupes:
        parts.append("4. Remove duplicate rows based on all data columns (keep first occurrence).")
    parts.append("5. Write the consolidated data to the current sheet starting at A1.")
    parts.append("6. Add a summary section: Source Files, Total Rows Before, Total Rows After, Duplicates Removed.")
    parts.append("Bold headers. Apply borders. Autofit columns.")
    parts.append("Start the data at A1. Use ws for the active sheet and wb for the workbook.")
    parts.append("In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.")
    parts.append("Respond with ONLY Python code. No explanation, no markdown, no backticks.")
    return "\n".join(parts)


def build_resume_parse_prompt(keywords, rank, extract_skills):
    parts = [
        "Create a resume parsing and analysis sheet.",
        f"Key Skills/Keywords: {keywords}",
        f"Rank Candidates: {'Yes' if rank else 'No'}",
        f"Extract Skills: {'Yes' if extract_skills else 'No'}",
        "",
        "Create a table with columns: Candidate #, Name, Email, Phone, "
        "Experience (Years), Current Role, Education, Key Skills Found.",
    ]
    if extract_skills:
        parts.append(
            "Add a 'Skills Extracted' column that lists all technical skills mentioned. "
            "Add a 'Skills Match %' column showing percentage of target keywords found."
        )
    if rank:
        parts.append(
            "Add a 'Rank' column based on: Skills Match (40%), Experience (30%), "
            "Education (20%), Keyword Density (10%). Sort by rank descending."
        )
    parts.append("Apply conditional formatting: Match > 80% = green, 50-80% = yellow, < 50% = red.")
    parts.append("Bold headers. Apply borders. Freeze top row.")
    parts.append("Start the data at A1. Use ws for the active sheet and wb for the workbook.")
    parts.append("In xlwings, a 1D list writes horizontally. Use nested 2D lists for rows.")
    parts.append("Respond with ONLY Python code. No explanation, no markdown, no backticks.")
    return "\n".join(parts)
