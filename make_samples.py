"""Generate a sample candidates.xlsx whose columns match the LOE template placeholders."""

from pathlib import Path

import pandas as pd

# Shared company / HR fields (same for every candidate in this demo)
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
}

candidates = [
    {
        "name": "John Tan Wei Ming",
        "first_name": "John",
        "nric": "S1234567A",
        "candidate_address": "12 Bedok Rise, #05-34, Singapore 469001",
        "job_title": "Senior Software Engineer",
        "department": "Engineering",
        "reporting_manager": "Daniel Lee",
        "reporting_manager_title": "Engineering Manager",
        "start_date": "1 July 2026",
        "probation_period": "three (3) months",
        "salary": "8,500",
        "bonus_target": "15%",
        "annual_leave": "18",
        "medical_leave": "14",
        "other_benefits": "Annual training budget of SGD 3,000",
        "reference_number": "ACME-LOE-2026-001",
        "offer_date": "19 June 2026",
        "acceptance_deadline": "26 June 2026",
    },
    {
        "name": "Mary Lim Hui Ling",
        "first_name": "Mary",
        "nric": "S7654321B",
        "candidate_address": "88 Toa Payoh Central, #11-02, Singapore 319901",
        "job_title": "Product Manager",
        "department": "Product",
        "reporting_manager": "Sarah Wong",
        "reporting_manager_title": "Head of Product",
        "start_date": "15 July 2026",
        "probation_period": "six (6) months",
        "salary": "9,200",
        "bonus_target": "20%",
        "annual_leave": "21",
        "medical_leave": "14",
        "other_benefits": "Flexible work arrangement, mobile phone allowance",
        "reference_number": "ACME-LOE-2026-002",
        "offer_date": "19 June 2026",
        "acceptance_deadline": "26 June 2026",
    },
    {
        "name": "Arjun Pillai",
        "first_name": "Arjun",
        "nric": "G9876543X",
        "candidate_address": "5 Holland Drive, #14-21, Singapore 271005",
        "job_title": "Data Analyst",
        "department": "Analytics",
        "reporting_manager": "Daniel Lee",
        "reporting_manager_title": "Engineering Manager",
        "start_date": "1 August 2026",
        "probation_period": "three (3) months",
        "salary": "6,000",
        "bonus_target": "10%",
        "annual_leave": "14",
        "medical_leave": "14",
        "other_benefits": "Professional certification sponsorship",
        "reference_number": "ACME-LOE-2026-003",
        "offer_date": "19 June 2026",
        "acceptance_deadline": "26 June 2026",
    },
]

rows = [{**company, **c} for c in candidates]
df = pd.DataFrame(rows)

out = Path(__file__).parent / "samples" / "candidates.xlsx"
out.parent.mkdir(parents=True, exist_ok=True)
df.to_excel(out, index=False)
print(f"Saved: {out}  ({len(df)} candidates, {len(df.columns)} columns)")
