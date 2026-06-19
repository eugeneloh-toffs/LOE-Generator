"""Generate samples/candidates_10.xlsx — 10 dummy candidates whose columns match
the template's placeholders exactly (columns are read from the template)."""

import io
from pathlib import Path

import pandas as pd
from docxtpl import DocxTemplate

HERE = Path(__file__).parent
tpl_bytes = (HERE / "samples" / "LOE_template.docx").read_bytes()
placeholders = sorted(DocxTemplate(io.BytesIO(tpl_bytes)).get_undeclared_template_variables())

# Fields shared by every candidate in this demo company
company = {
    "company_name": "Acme Technologies Pte Ltd",
    "company_address": "1 Marina Boulevard, #20-01, Singapore 018989",
    "company_uen": "201912345A",
    "company_phone": "+65 6123 4567",
    "company_email": "hr@acmetech.com.sg",
    "hr_name": "Rachel Ng",
    "hr_title": "Head of Human Resources",
    "work_location": "1 Marina Boulevard, Singapore",
    "working_hours": "9:00am to 6:00pm",
    "working_days": "Monday to Friday",
    "lunch_break": "one-hour",
    "payment_date": "the last working day",
    "probation_notice": "one (1) week's",
    "probation_extension": "three (3) months",
    "notice_period": "one (1) month's",
    "non_compete_period": "six (6) months",
    "non_compete_geography": "Singapore",
    "lookback_period": "twelve (12) months",
    "medical_benefits": "Outpatient and specialist coverage under the Company group scheme",
    "insurance_coverage": "Group term life and personal accident insurance",
    "hospitalisation_leave": "60",
    "work_pass_type": "Employment Pass or Singapore citizenship/PR",
    "offer_date": "19 June 2026",
    "acceptance_deadline": "26 June 2026",
}

# 10 varied candidates
people = [
    ("Tan Wei Ming", "John", "S1234567A", "12 Bedok Rise, #05-34, Singapore 469001",
     "Senior Software Engineer", "Engineering", "Daniel Lee", "Engineering Manager",
     "1 July 2026", "three (3) months", "8,500", "15%", "18", "14",
     "Annual training budget of SGD 3,000"),
    ("Lim Hui Ling", "Mary", "S7654321B", "88 Toa Payoh Central, #11-02, Singapore 319901",
     "Product Manager", "Product", "Sarah Wong", "Head of Product",
     "15 July 2026", "six (6) months", "9,200", "20%", "21", "14",
     "Flexible work arrangement, mobile phone allowance"),
    ("Arjun Pillai", "Arjun", "G9876543X", "5 Holland Drive, #14-21, Singapore 271005",
     "Data Analyst", "Analytics", "Daniel Lee", "Engineering Manager",
     "1 August 2026", "three (3) months", "6,000", "10%", "14", "14",
     "Professional certification sponsorship"),
    ("Nurul Aisyah Binte Rahman", "Nurul", "S8123456C", "3 Tampines Street 32, #09-11, Singapore 529281",
     "UX Designer", "Design", "Priya Menon", "Design Lead",
     "1 July 2026", "three (3) months", "6,800", "10%", "16", "14",
     "Home office setup allowance of SGD 1,500"),
    ("Chen Jia Hao", "Jia Hao", "S9234567D", "20 Clementi Avenue 4, #07-08, Singapore 129908",
     "DevOps Engineer", "Engineering", "Daniel Lee", "Engineering Manager",
     "15 August 2026", "three (3) months", "9,000", "15%", "18", "14",
     "On-call allowance, annual training budget of SGD 3,000"),
    ("Siti Khadijah", "Siti", "S8345678E", "45 Jurong West Street 41, #12-345, Singapore 640045",
     "HR Business Partner", "Human Resources", "Rachel Ng", "Head of Human Resources",
     "1 September 2026", "six (6) months", "7,200", "12%", "18", "14",
     "Professional membership fees reimbursement"),
    ("Rajesh Kumar", "Rajesh", "G8456789Y", "7 Serangoon North Avenue 5, #03-19, Singapore 554912",
     "Account Executive", "Sales", "Michael Tan", "Sales Director",
     "1 July 2026", "three (3) months", "5,500", "30%", "14", "14",
     "Sales commission scheme, transport allowance"),
    ("Wong Mei Fang", "Mei Fang", "S7567890F", "101 Bukit Timah Road, #18-02, Singapore 229899",
     "Finance Manager", "Finance", "Lawrence Goh", "Chief Financial Officer",
     "1 October 2026", "six (6) months", "10,500", "20%", "21", "14",
     "Car allowance of SGD 1,200 per month"),
    ("David Ong Kang Wei", "David", "S9678901G", "33 Pasir Ris Drive 6, #06-77, Singapore 519421",
     "Marketing Specialist", "Marketing", "Sarah Wong", "Head of Product",
     "15 July 2026", "three (3) months", "5,800", "10%", "16", "14",
     "Mobile phone allowance"),
    ("Foo Li Wei", "Li Wei", "S8789012H", "9 Ang Mo Kio Avenue 3, #15-201, Singapore 569933",
     "QA Engineer", "Engineering", "Daniel Lee", "Engineering Manager",
     "1 August 2026", "three (3) months", "6,200", "10%", "16", "14",
     "Annual training budget of SGD 3,000"),
]

rows = []
for idx, (name, first, nric, addr, title, dept, mgr, mgr_title, start, prob,
          salary, bonus, al, ml, other) in enumerate(people, start=1):
    row = {
        **company,
        "name": name,
        "first_name": first,
        "nric": nric,
        "candidate_address": addr,
        "job_title": title,
        "department": dept,
        "reporting_manager": mgr,
        "reporting_manager_title": mgr_title,
        "start_date": start,
        "probation_period": prob,
        "salary": salary,
        "bonus_target": bonus,
        "annual_leave": al,
        "medical_leave": ml,
        "other_benefits": other,
        "reference_number": f"ACME-LOE-2026-{idx:03d}",
    }
    rows.append(row)

df = pd.DataFrame(rows)

# Order columns: template placeholders first (sorted), then any extras
ordered = [c for c in placeholders if c in df.columns] + [c for c in df.columns if c not in placeholders]
df = df[ordered]

# Sanity check against template
missing = set(placeholders) - set(df.columns)
assert not missing, f"Missing columns for placeholders: {missing}"

out = HERE / "samples" / "candidates_10.xlsx"
df.to_excel(out, index=False)
print(f"Saved: {out}")
print(f"Rows: {len(df)} | Columns: {len(df.columns)} | All {len(placeholders)} placeholders covered: {not missing}")
