"""
SOA Completion Agent — Python Backend
Brightday Australia
Version: 1.2

Reads:  Fact Finder (.xlsx)  →  Fact Finder tab (via python-calamine)
        SOA Template (.docx) →  find & replace {{codes}} (via python-docx)

Outputs: Completed SOA draft (.docx) with all insertions in red font.
         Unmapped codes are left as raw {{code}} text.

Dependencies:
    pip install flask python-docx python-calamine gunicorn gevent

Run locally:
    python app.py
    Open http://localhost:5000

Deploy to Render:
    Start Command: gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --timeout 120
    Environment vars: SECRET_KEY, USERS (see DEPLOYMENT_GUIDE.md)
"""

from flask import Flask, request, send_file, jsonify, session, redirect, url_for, render_template_string
from docx import Document
from docx.shared import RGBColor, Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import date
import io
import traceback
import re
import copy
import os
import hashlib

app = Flask(__name__)

# ─────────────────────────────────────────────
# AUTH CONFIG
# Read credentials from environment variables.
# Set these in Render dashboard — never hardcode.
# ─────────────────────────────────────────────
app.secret_key = os.environ.get("SECRET_KEY", "change-this-in-production")

# USERS dict — username: hashed password
# To generate a hash: python3 -c "import hashlib; print(hashlib.sha256('yourpassword'.encode()).hexdigest())"
# Add as many users as needed in the USERS env var format:
#   USERS=username1:hash1,username2:hash2
def load_users():
    users_env = os.environ.get("USERS", "")
    users = {}
    for entry in users_env.split(","):
        entry = entry.strip()
        if ":" in entry:
            username, pw_hash = entry.split(":", 1)
            users[username.strip().lower()] = pw_hash.strip()
    return users

def check_password(username, password):
    users = load_users()
    pw_hash = hashlib.sha256(password.encode()).hexdigest()
    return users.get(username.lower()) == pw_hash

def logged_in():
    return session.get("authenticated") is True

# ─────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────
RED   = RGBColor(0xFF, 0x00, 0x00)
BLACK = RGBColor(0x00, 0x00, 0x00)

# Fact Finder tab — column letters for multi-fund fields (up to 5 funds)
# Columns: B=2, D=4, F=6, H=8, J=10
FUND_COLS = [2, 4, 6, 8, 10]


# ─────────────────────────────────────────────
# FACT FINDER READER
# ─────────────────────────────────────────────

def read_fact_finder(xlsx_bytes, risk_profile, no_insurance_flag,
                     no_trauma_flag=False, no_salsac_flag=False,
                     insurance_only_flag=False, scenario=""):
    """
    Read the Fact Finder xlsx and return:
        - data dict  { "{{CODE}}": "value" }
        - conditionals dict  { "DELETE_KEY": True/False }

    UI flags:
        no_insurance_flag    -> drives {{DeleteIfNoInsuranceAtAll}} and {{DeleteIfNoInsuranceAdvice}}
        no_trauma_flag       -> drives {{DeleteIfNoScopedTrauma}}
        no_salsac_flag       -> drives {{DeleteIfNoSalarySacrificeAdvice}}
        insurance_only_flag  -> drives {{DeleteIfInsuranceOnlyClient}}
        scenario             -> "1".."6"; keeps {{ScenarioN}} block, deletes the other five
    """
    from python_calamine import CalamineWorkbook
    cal_wb   = CalamineWorkbook.from_filelike(io.BytesIO(xlsx_bytes))
    sheet    = cal_wb.get_sheet_by_name("Fact Finder")
    raw_rows = sheet.to_python(skip_empty_area=False)

    # Build a row/col lookup dict identical to openpyxl's ws.cell(row, col).value
    # calamine returns a list of rows; each row is a list of values
    # rows and cols are 1-indexed to match existing code
    cell_data = {}
    for r_idx, row in enumerate(raw_rows, start=1):
        for c_idx, val in enumerate(row, start=1):
            cell_data[(r_idx, c_idx)] = val

    def cell(row, col):
        """Return cleaned string value from a cell, or '' if empty/zero placeholder.
        Per spec: a numeric 0 (int or float) is treated as empty."""
        v = cell_data.get((row, col))
        if v is None:
            return ""
        # Numeric zero -> empty (covers calamine returning 0.0 for blank-looking cells)
        if isinstance(v, (int, float)) and not isinstance(v, bool) and float(v) == 0:
            return ""
        s = str(v).strip()
        if s in ("", "0", "0.0", "00:00:00", "#REF!", "None"):
            return ""
        return s

    def cells_across(row, cols=FUND_COLS):
        """Return list of non-empty values across multiple fund columns."""
        return [cell(row, c) for c in cols if cell(row, c)]

    def join_funds(row, sep=", "):
        vals = cells_across(row)
        return sep.join(vals) if vals else ""

    def sum_funds(row):
        total = 0
        for c in FUND_COLS:
            v = cell_data.get((row, c))
            try:
                total += float(str(v).replace(",", "").replace("$", ""))
            except Exception:
                pass
        return total

    def currency(row, col):
        """Currency-format a cell value. Returns '' for empty cells AND numeric 0."""
        v = cell_data.get((row, col))
        if v is None:
            return ""
        try:
            n = float(str(v).replace(',', '').replace('$', ''))
        except (ValueError, TypeError):
            return ""
        if n == 0:
            return ""
        return f"${n:,.0f}"

    def currency_sum(row):
        s = sum_funds(row)
        return f"${s:,.0f}" if s else ""

    def age_from_dob(row, col=2):
        v = cell_data.get((row, col))
        if not v:
            return None
        try:
            from datetime import datetime
            if hasattr(v, 'year'):
                dob = v
            else:
                for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
                    try:
                        dob = datetime.strptime(str(v), fmt)
                        break
                    except ValueError:
                        continue
                else:
                    return None
            today = date.today()
            return today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))
        except Exception:
            return None

    def format_au_phone(raw):
        """Format an Australian phone number as '0XXX XXX XXX'.
        Handles inputs with spaces, dashes, +61 prefix, or already-formatted.
        Returns the original string if it doesn't look like a 10-digit AU number."""
        if not raw:
            return ""
        digits = "".join(c for c in str(raw) if c.isdigit())
        # Convert +61... or 61... (11 digits) to 0...
        if len(digits) == 11 and digits.startswith("61"):
            digits = "0" + digits[2:]
        if len(digits) == 10 and digits.startswith("0"):
            return f"{digits[:4]} {digits[4:7]} {digits[7:]}"
        return str(raw).strip()

    def format_date(row, col=2):
        v = cell_data.get((row, col))
        if not v:
            return ""
        try:
            if hasattr(v, 'strftime'):
                return v.strftime("%d/%m/%Y")
            return str(v)
        except Exception:
            return str(v)

    # ── Personal Details ──
    title       = cell(10, 2)
    first_name  = cell(11, 2)
    last_name   = cell(13, 2)
    # Per spec: ClientFullName = first + last only (middle name skipped)
    full_name_parts = [p for p in [first_name, last_name] if p]
    full_name   = " ".join(full_name_parts)
    dob_str     = format_date(15, 2)
    age         = age_from_dob(15, 2)
    phone       = format_au_phone(cell(16, 2))
    email       = cell(17, 2)
    address_parts = [p for p in [cell(18,2), cell(19,2), cell(20,2), cell(21,2)] if p]
    address     = ", ".join(address_parts)

    # ── Employment ──
    occupation  = cell(28, 2)
    emp_status  = cell(23, 2)

    # ── Income ──
    gross_income_raw = cell_data.get((32, 2))
    try:
        gross_income_num = float(str(gross_income_raw).replace(",","").replace("$",""))
        gross_income = f"${gross_income_num:,.0f}"
    except Exception:
        gross_income_num = 0
        gross_income = ""

    sgc_pct_raw = cell_data.get((34, 2))
    try:
        sgc_pct = float(str(sgc_pct_raw).replace("%","")) / 100
    except Exception:
        sgc_pct = 0.12
    super_contribution = f"${gross_income_num * sgc_pct:,.0f}" if gross_income_num else ""

    salary_sacrifice_raw = cell_data.get((35, 2))
    try:
        salary_sacrifice = f"${float(str(salary_sacrifice_raw).replace(',','').replace('$','')):,.0f}"
        annualised_salary_sacrifice = salary_sacrifice
    except Exception:
        salary_sacrifice = ""
        annualised_salary_sacrifice = ""

    # ── Retirement Age ──
    # Per spec: default to B8; use B9 only if B9 is non-empty AND differs from B8.
    # cell() returns "" for both blank and numeric 0, so the simple check covers both.
    ret_age_1 = cell(8, 2)
    ret_age_2 = cell(9, 2)
    retirement_age = ret_age_2 if (ret_age_2 and ret_age_2 != ret_age_1) else ret_age_1

    # ── Spouse ──
    spouse_dob  = format_date(47, 2)
    spouse_income_raw = cell_data.get((49, 2))
    try:
        spouse_income = f"${float(str(spouse_income_raw).replace(',','').replace('$','')):,.0f}"
    except Exception:
        spouse_income = ""
    spouse_balance_raw = cell_data.get((50, 2))
    try:
        spouse_balance = f"${float(str(spouse_balance_raw).replace(',','').replace('$','')):,.0f}"
    except Exception:
        spouse_balance = ""

    # ── Dependants ──
    has_spouse = bool(cell(46, 2))  # Spouse Name row
    dep_ages = cells_across(56)
    no_dependants = (1 if has_spouse else 0) + len(dep_ages)

    # ── Assets & Liabilities ──
    primary_residence_val = currency(73, 2)
    primary_residence_debt = currency(74, 2)
    investment_prop_val = currency(76, 2)
    investment_prop_debt = currency(77, 2)
    other_asset1_val = currency(79, 2)
    personal_loan1_val = currency(81, 2)

    # Total assets
    total_assets = 0
    for r, c in [(73,2),(76,2),(79,2)]:
        v = cell_data.get((r, c))
        try:
            total_assets += float(str(v).replace(",","").replace("$",""))
        except Exception:
            pass
    # Add super balances
    total_assets += sum_funds(94)
    total_assets_str = f"${total_assets:,.0f}" if total_assets else ""

    total_liabilities = 0
    for r, c in [(74,2),(77,2),(81,2)]:
        v = cell_data.get((r, c))
        try:
            total_liabilities += float(str(v).replace(",","").replace("$",""))
        except Exception:
            pass
    total_liabilities_str = f"${total_liabilities:,.0f}" if total_liabilities else ""

    # ── Super Funds ──
    current_super_funds = join_funds(92)
    current_super_balance = currency_sum(94)
    current_balance = current_super_balance

    # ── Insurance across funds ──
    def insurance_across(row):
        vals = []
        for c in FUND_COLS:
            v = cell_data.get((row, c))
            if v:
                s = str(v).strip()
                if s not in ("0","","None"):
                    vals.append(s)
        if not vals:
            return ""
        unique = list(dict.fromkeys(vals))
        return " / ".join(unique)

    life_ins   = insurance_across(102)
    tpd_ins    = insurance_across(103)
    ip_month   = insurance_across(104)
    ip_wait    = insurance_across(105)
    ip_benefit = insurance_across(106)
    premiums   = insurance_across(107)

    # ── Binding Death Nominee ──
    # Per spec: read from row 63 across fund columns (was row 62).
    # cell() now treats numeric 0 / blank as empty so zero-cells are skipped.
    nominee_names = [cell(63, c) for c in FUND_COLS if cell(63, c)]
    binding_death_nominee = ", ".join(nominee_names) if nominee_names else "N/A"

    # ── Current Date ──
    current_date = date.today().strftime("%d %B %Y")

    # ── Risk Profile (from UI selection) ──
    current_risk_profile = risk_profile  # passed in from form

    # ─────────────────────────────────
    # BUILD DATA DICT
    # ─────────────────────────────────
    data = {
        "{{Title}}":                             title,
        "{{ClientFullName}}":                    full_name,
        "{{ClientFirstName}}":                   first_name,
        "{{ClientLastName}}":                    last_name,
        "{{ClientDOB}}":                         dob_str,
        "{{ClientAddress}}":                     address,
        "{{ClientPhone}}":                       phone,
        "{{ClientEmail}}":                       email,
        "{{ClientOccupation}}":                  occupation,
        "{{ClientSalary}}":                      gross_income,
        "{{fld_SuperContribution}}":             super_contribution,
        "{{fld_SalarySacrifice}}":               salary_sacrifice,
        "{{CurrentSuperFunds}}":                 current_super_funds,
        "{{SpouseDOB}}":                         spouse_dob,
        "{{SpouseIncome}}":                      spouse_income,
        "{{SpouseBalance}}":                     spouse_balance,
        "{{NoDependants}}":                      str(no_dependants),
        "{{fld_CurrentSuperannuationBalance}}":  current_super_balance,
        "{{CurrentLifeInsurance}}":              life_ins,
        "{{CurrentTPDInsurance}}":               tpd_ins,
        "{{CurrentIncomeProtectionPerMonth}}":   ip_month,
        "{{CurrentIncomeProtectionWaitingPeriod}}": ip_wait,
        "{{CurrentIncomeProtectionBenefitPeriod}}": ip_benefit,
        "{{CurrentSuperPremiums}}":              premiums,
        "{{ValueOfPrimaryResidence}}":           primary_residence_val,
        "{{DebtOnPrimaryResidence}}":            primary_residence_debt,
        "{{ValueOfInvestmentProperty}}":         investment_prop_val,
        "{{DebtOnInvestmentProperty}}":          investment_prop_debt,
        "{{OtherAsset1Value}}":                  other_asset1_val,
        "{{PersonalLoan1Value}}":                personal_loan1_val,
        "{{TotalAssetValue}}":                   total_assets_str,
        "{{TotalLiabilityValue}}":               total_liabilities_str,
        "{{RetirementAge}}":                     retirement_age,
        "{{CurrentBalance}}":                    current_balance,
        "{{CurrentAge}}":                        str(age) if age else "",
        "{{CurrentDate}}":                       current_date,
        "{{AnnualisedSalarySacrificeAmount}}":   annualised_salary_sacrifice,
        "{{BindingDeathNominee}}":               binding_death_nominee,
        "{{CurrentRiskProfile}}":                current_risk_profile,
        "{{EmploymentStatus}}":                  emp_status,
        "{{MaritalStatus}}":                     cell(45, 2),
        # Goals — left as raw codes (adviser-completed)
        # Table codes — left as raw codes (adviser-completed)
        # $r1-$r4, $p1-$p4 — adviser-completed (different placeholder syntax, regex ignores them)
    }

    # ─────────────────────────────────
    # CONDITIONALS
    # ─────────────────────────────────
    total_balance = sum_funds(94)
    super_contribution_num = gross_income_num * sgc_pct if gross_income_num else 0

    # Row 100: insurance in fund — check all fund columns
    has_any_insurance = any(
        str(cell_data.get((100, c)) or "").strip().lower() == "yes"
        for c in FUND_COLS
    )

    # Row 108: medically underwritten
    has_underwritten = any(
        str(cell_data.get((108, c)) or "").strip().lower() == "medically underwritten"
        for c in FUND_COLS
    )

    conditionals = {
        # True = DELETE this block
        "DeleteIfAgeGreaterThan55":              (age is not None and age >= 55),
        "DeleteIfAgeLessThan55":                 (age is not None and age < 55),
        "DeleteIfBalanceBelow500k":              (total_balance < 500_000),
        "DeleteIfSuperContributionsBelow30k":    (super_contribution_num < 30_000),
        "DeleteIfNoCurrentInsurance":            (not has_any_insurance),
        "DeleteIfNoInsuranceAtAll":              no_insurance_flag,        # legacy UI checkbox
        "DeleteIfNoInsuranceAdvice":             no_insurance_flag,        # new spec name, same checkbox
        "DeleteIfNoScopedTrauma":                no_trauma_flag,           # UI checkbox
        "DeleteIfNoSalarySacrificeAdvice":       no_salsac_flag,           # UI checkbox
        "DeleteIfInsuranceOnlyClient":           insurance_only_flag,      # UI checkbox
        # Per Code Map: delete the no-scoped-insurance block when adviser picks scenario 5 or 6
        "DeleteIfNoScopedInsurance":             (scenario in {"5", "6"}),
        "DeleteIfNoTrauma":                      False,  # unmapped — never delete
        # Per Code Map: delete the personal-deductible-contributions block when both
        # B35 (salary sacrifice) AND B36 are non-empty / non-zero (cell() treats 0 as empty)
        "DeleteIfPersonalDeductibleContributions": bool(cell(35, 2) and cell(36, 2)),
        # Per Code Map: total liabilities = B74 + B77 + B81 (computed above as total_liabilities)
        "DeleteIfNoDebts":                       (total_liabilities == 0),
        "DeleteIfClientHasDebts":                (total_liabilities > 0),
        "DeleteifNoCurrentUnderwrittenInsurance": (not has_underwritten),
    }

    # Scenarios 1–6: keep selected, delete the other five.
    # If no scenario selected, all six are kept (markers stripped).
    for n in range(1, 7):
        conditionals[f"DeleteScenario{n}"] = (scenario != "" and scenario != str(n))

    return data, conditionals


# ─────────────────────────────────────────────
# SOA DOCUMENT PROCESSOR
# ─────────────────────────────────────────────

# Codes that are intentionally left as raw {{code}} — never replaced
UNMAPPED_CODES = {
    "{{Date}}",
    "{{OtherAsset1}}",
    "{{OtherAsset2}}",
    "{{OtherAsset2Value}}",
    "{{PersonalLoan2Value}}",
    "{{PersonalLoan1}}",
    "{{PersonalLoan2}}",
    "{{NeedsAnalysisLifeInsurance}}",
    "{{NeedsAnalysisTPD}}",
    "{{NeedsAnalysisIP}}",
    "{{NeedsAnalysisTrauma}}",
    "{{Tbl_SalarySacrifice}}",
    "{{tbl_CurrentSuperFundsRiskProfilePerformance}}",
    "{{Make personal deductible contributions/Salary sacrifice}}",
    "{{DeleteIfNoTrauma}}",
    "{{EndDeleteIfNoTrauma}}",
    "{{CurrentInsuer}}",
    "{{SalarySacrificeAmount}}",
    "{{SalarySacrificeFrequency}}",
    "{{NetTaxSavings}}",
    "{{zzz}}",
    "{{SuperGoal}}",
    "{{InsuranceGoal}}",
    "{{SalarySacrificeGoal}}",
    "{{EstatePlanningGoal}}",
    "{{RetirementGoal}}",
    "{{DeleteIfNoInsuranceAtAll}}",
    "{{EndDeleteIfNoInsuranceAtAll}}",
}

# Pair up conditional block tags
CONDITIONAL_PAIRS = [
    ("{{DeleteIfAgeGreaterThan55}}",              "{{EndDeleteIfAgeGreaterThan55}}",              "DeleteIfAgeGreaterThan55"),
    ("{{DeleteIfAgeLessThan55}}",                 "{{EndDeleteIfAgeLessThan55}}",                 "DeleteIfAgeLessThan55"),
    ("{{DeleteIfBalanceBelow500k}}",              "{{EndDeleteIfBalanceBelow500k}}",              "DeleteIfBalanceBelow500k"),
    ("{{DeleteIfSuperContributionsBelow30k}}",    "{{EndDeleteIfSuperContributionsBelow30k}}",    "DeleteIfSuperContributionsBelow30k"),
    ("{{DeleteIfNoCurrentInsurance}}",            "{{EndDeleteIfNoCurrentInsurance}}",            "DeleteIfNoCurrentInsurance"),
    ("{{DeleteifNoCurrentUnderwrittenInsurance}}","{{EndDeleteifNoCurrentUnderwrittenInsurance}}","DeleteifNoCurrentUnderwrittenInsurance"),
    # New spec — driven by UI checkboxes
    ("{{DeleteIfNoInsuranceAdvice}}",             "{{EndDeleteIfNoInsuranceAdvice}}",             "DeleteIfNoInsuranceAdvice"),
    ("{{DeleteIfNoScopedTrauma}}",                "{{EndDeleteIfNoScopedTrauma}}",                "DeleteIfNoScopedTrauma"),
    ("{{DeleteIfNoSalarySacrificeAdvice}}",       "{{EndDeleteIfNoSalarySacrificeAdvice}}",       "DeleteIfNoSalarySacrificeAdvice"),
    ("{{DeleteIfInsuranceOnlyClient}}",           "{{EndDeleteIfInsuranceOnlyClient}}",           "DeleteIfInsuranceOnlyClient"),
    # Auto from FF / scenario selection
    ("{{DeleteIfNoScopedInsurance}}",             "{{EndDeleteIfNoScopedInsurance}}",             "DeleteIfNoScopedInsurance"),
    ("{{DeleteIfNoDebts}}",                       "{{EndDeleteIfNoDebts}}",                       "DeleteIfNoDebts"),
    ("{{DeleteIfClientHasDebts}}",                "{{EndDeleteIfClientHasDebts}}",                "DeleteIfClientHasDebts"),
    ("{{DeleteIfPersonalDeductibleContributions}}","{{EndDeleteIfPersonalDeductibleContributions}}","DeleteIfPersonalDeductibleContributions"),
    # Scenario 1–6 — adviser picks one in the UI; the other five are deleted
    ("{{Scenario1}}", "{{EndScenario1}}", "DeleteScenario1"),
    ("{{Scenario2}}", "{{EndScenario2}}", "DeleteScenario2"),
    ("{{Scenario3}}", "{{EndScenario3}}", "DeleteScenario3"),
    ("{{Scenario4}}", "{{EndScenario4}}", "DeleteScenario4"),
    ("{{Scenario5}}", "{{EndScenario5}}", "DeleteScenario5"),
    ("{{Scenario6}}", "{{EndScenario6}}", "DeleteScenario6"),
]

# For DeleteIfNoInsuranceAtAll — single tag (no end tag), marks start of section to delete
# We treat the content following it until next section heading as the block
NO_INSURANCE_SINGLE_TAG = "{{DeleteIfNoInsuranceAtAll}}"


def get_full_text(paragraph):
    return "".join(run.text for run in paragraph.runs)


def para_contains(paragraph, code):
    return code in get_full_text(paragraph)


def replace_code_in_run(run, code, value, use_red):
    """Replace a code in a single run, applying red font to the replacement."""
    if code not in run.text:
        return
    parts = run.text.split(code)
    # If only one part before and after — simple case
    if len(parts) == 2:
        before, after = parts
        run.text = before
        # Insert red replacement run after this run
        p = run._r.getparent()
        idx = list(p).index(run._r)

        def make_run(text, red):
            from docx.oxml import OxmlElement
            r_el = OxmlElement('w:r')
            # Copy rPr from original run
            if run._r.find(qn('w:rPr')) is not None:
                rPr = copy.deepcopy(run._r.find(qn('w:rPr')))
                # Set or remove colour
                color_el = rPr.find(qn('w:color'))
                if color_el is None:
                    color_el = OxmlElement('w:color')
                    rPr.append(color_el)
                if red:
                    color_el.set(qn('w:val'), 'FF0000')
                else:
                    color_el.set(qn('w:val'), 'auto')
                r_el.append(rPr)
            else:
                if red:
                    rPr = OxmlElement('w:rPr')
                    color_el = OxmlElement('w:color')
                    color_el.set(qn('w:val'), 'FF0000')
                    rPr.append(color_el)
                    r_el.append(rPr)
            t_el = OxmlElement('w:t')
            t_el.text = text
            if text.startswith(' ') or text.endswith(' '):
                t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            r_el.append(t_el)
            return r_el

        # Insert replacement
        if value:
            p.insert(idx + 1, make_run(value, use_red))
        # Insert after-text
        if after:
            p.insert(idx + 2, make_run(after, False))
    else:
        # Multiple occurrences in one run — replace all
        new_text = run.text.replace(code, value)
        run.text = new_text
        if use_red and value:
            run.font.color.rgb = RED


def process_paragraph_text(paragraph, data, unmapped):
    """Replace all known codes in a paragraph. Leave unmapped codes untouched."""
    full = get_full_text(paragraph)
    if "{{" not in full:
        return

    # Find all codes in this paragraph
    codes_present = re.findall(r'\{\{[^}]+\}\}', full)

    for code in codes_present:
        if code in unmapped:
            continue  # Leave raw
        if code in data:
            value = data[code]
            # Work run by run
            for run in paragraph.runs:
                if code in run.text:
                    replace_code_in_run(run, code, value, use_red=True)


def process_table(table, data, unmapped):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                process_paragraph_text(paragraph, data, unmapped)
            for nested_table in cell.tables:
                process_table(nested_table, data, unmapped)


def collect_all_paragraphs(doc):
    """Return flat list of (paragraph, parent_element, index) for body + tables."""
    items = []
    body = doc.element.body
    for i, child in enumerate(body):
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            from docx.text.paragraph import Paragraph
            items.append(Paragraph(child, doc))
        elif tag == 'tbl':
            from docx.table import Table
            tbl = Table(child, doc)
            for row in tbl.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        items.append(p)
    return items


def apply_conditional_deletions(doc, conditionals):
    """
    Walk through document body elements.
    When a start-tag paragraph is found and its condition is True,
    collect and remove all elements up to and including the end-tag paragraph.
    """
    body = doc.element.body
    elements = list(body)

    def get_para_text(el):
        return "".join(t.text or "" for t in el.iter(qn('w:t')))

    for start_tag, end_tag, condition_key in CONDITIONAL_PAIRS:
        should_delete = conditionals.get(condition_key, False)
        if not should_delete:
            # Still remove the marker tags themselves (they're not content)
            to_remove = []
            for el in list(body):
                txt = get_para_text(el)
                if start_tag in txt or end_tag in txt:
                    to_remove.append(el)
            for el in to_remove:
                body.remove(el)
            continue

        # Delete everything between (and including) start and end tags
        in_block = False
        to_remove = []
        for el in list(body):
            txt = get_para_text(el)
            if start_tag in txt:
                in_block = True
            if in_block:
                to_remove.append(el)
            if end_tag in txt and in_block:
                in_block = False
        for el in to_remove:
            try:
                body.remove(el)
            except ValueError:
                pass

    # Handle DeleteIfNoInsuranceAtAll (no end tag)
    # Remove the single marker tag paragraph regardless
    should_delete_no_ins = conditionals.get("DeleteIfNoInsuranceAtAll", False)
    to_remove = []
    in_block = False
    for el in list(body):
        txt = get_para_text(el)
        if NO_INSURANCE_SINGLE_TAG in txt:
            to_remove.append(el)  # always remove the tag itself
            if should_delete_no_ins:
                in_block = True
            continue
        if in_block:
            # Delete until we hit the next heading-level paragraph or end of section
            # Heuristic: stop at next paragraph that has bold text > 12pt or is a heading style
            tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
            if tag == 'p':
                style = el.find('.//' + qn('w:pStyle'))
                style_val = style.get(qn('w:val'), '') if style is not None else ''
                if 'Heading' in style_val or style_val.startswith('h'):
                    in_block = False
                    continue
            to_remove.append(el)
    for el in to_remove:
        try:
            body.remove(el)
        except ValueError:
            pass


def process_soa(template_bytes, data, conditionals):
    """Main processor — returns completed docx as bytes."""
    import gc

    # Load document then immediately free the raw bytes from memory
    buf = io.BytesIO(template_bytes)
    doc = Document(buf)
    del template_bytes, buf
    gc.collect()

    # Step 1: Apply conditional block deletions
    apply_conditional_deletions(doc, conditionals)

    # Step 2: Replace codes in body paragraphs
    for paragraph in doc.paragraphs:
        process_paragraph_text(paragraph, data, UNMAPPED_CODES)

    # Step 3: Replace codes in tables
    for table in doc.tables:
        process_table(table, data, UNMAPPED_CODES)

    # Step 4: Replace codes in headers and footers
    for section in doc.sections:
        for hdr in [section.header, section.footer,
                    section.even_page_header, section.even_page_footer,
                    section.first_page_header, section.first_page_footer]:
            if hdr:
                for paragraph in hdr.paragraphs:
                    process_paragraph_text(paragraph, data, UNMAPPED_CODES)
                for table in hdr.tables:
                    process_table(table, data, UNMAPPED_CODES)

    out = io.BytesIO()
    doc.save(out)
    del doc
    gc.collect()
    out.seek(0)
    return out


# ─────────────────────────────────────────────
# LOGIN PAGE HTML
# ─────────────────────────────────────────────

LOGIN_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SOA Agent — Brightday Australia</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Ubuntu:wght@300;400;500;700&display=swap" rel="stylesheet">
<link href="https://api.fontshare.com/v2/css?f[]=general-sans@300,400,500,600,700&display=swap" rel="stylesheet">
<style>
  :root {
    --navy:        #123559;
    --navy-deep:   #0C243D;
    --raspberry:   #F50D74;
    --raspberry-light: #FF4593;
    --bg:          #FFFFFF;
    --surface:     #FFFFFF;
    --surface-soft:#F7F8FA;
    --text:        #0E2640;
    --text-muted:  #5C6B7E;
    --text-soft:   #8A95A4;
    --border:      #E5E9EF;
    --border-strong:#CBD2DB;
    --red:         #DA2929;
    --red-soft:    #FCEEEE;
  }
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  html, body { height: 100%; }
  body {
    background: var(--bg);
    color: var(--text);
    font-family: 'General Sans', 'Ubuntu', system-ui, sans-serif;
    font-weight: 400;
    min-height: 100vh;
    display: flex; flex-direction: column;
  }
  body::before {
    content: ''; position: fixed; top: -180px; right: -180px;
    width: 520px; height: 520px;
    background: radial-gradient(circle, rgba(245,13,116,0.10) 0%, transparent 70%);
    pointer-events: none; z-index: 0;
  }
  body::after {
    content: ''; position: fixed; bottom: -200px; left: -200px;
    width: 520px; height: 520px;
    background: radial-gradient(circle, rgba(18,53,89,0.06) 0%, transparent 70%);
    pointer-events: none; z-index: 0;
  }
  .site-header {
    background: var(--navy);
    border-bottom: 3px solid var(--raspberry);
    z-index: 1; position: relative;
  }
  .header-inner {
    max-width: 1080px; margin: 0 auto;
    padding: 18px 32px;
    display: flex; align-items: center; gap: 14px;
  }
  .logo-mark {
    width: 38px; height: 38px;
    display: flex; align-items: center; justify-content: center; flex-shrink: 0;
  }
  .logo-mark svg { width: 100%; height: 100%; display: block; }
  .header-text { display: flex; flex-direction: column; line-height: 1; }
  .header-text .brand {
    font-family: 'Ubuntu', sans-serif; font-size: 22px; font-weight: 700;
    color: #FFFFFF; letter-spacing: -0.5px;
  }
  .header-text .sub {
    font-family: 'General Sans', sans-serif;
    font-size: 10px; color: var(--raspberry-light); letter-spacing: 2.2px; font-weight: 500;
    text-transform: uppercase; margin-top: 5px;
  }

  .login-wrap {
    flex: 1;
    display: flex; align-items: center; justify-content: center;
    padding: 48px 24px;
    position: relative; z-index: 1;
  }
  .card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 44px 40px;
    width: 100%; max-width: 420px;
    animation: fadeUp 0.6s ease both;
    box-shadow: 0 12px 40px rgba(14,38,64,0.08);
  }
  h1 {
    font-family: 'Ubuntu', sans-serif; font-size: 28px; font-weight: 700;
    color: var(--navy);
    line-height: 1.1; margin-bottom: 8px; letter-spacing: -0.6px;
  }
  h1 em { font-style: normal; color: var(--raspberry); }
  .subtitle {
    font-family: 'General Sans', sans-serif;
    font-size: 13px; color: var(--text-muted); margin-bottom: 32px; line-height: 1.6;
  }
  .field { margin-bottom: 18px; }
  .field label {
    font-family: 'General Sans', sans-serif;
    display: block; font-size: 10px; font-weight: 600;
    letter-spacing: 2px; text-transform: uppercase;
    color: var(--text-muted); margin-bottom: 6px;
  }
  .field input {
    width: 100%;
    background: var(--surface);
    border: 1px solid var(--border);
    color: var(--text);
    font-family: 'General Sans', sans-serif; font-size: 13px;
    padding: 12px 14px; outline: none;
    border-radius: 8px;
    transition: border-color 0.2s, box-shadow 0.2s;
  }
  .field input:hover { border-color: var(--border-strong); }
  .field input:focus { border-color: var(--raspberry); box-shadow: 0 0 0 3px rgba(245,13,116,0.10); }
  .btn-login {
    width: 100%;
    background: var(--raspberry); color: #FFFFFF; border: none;
    padding: 14px;
    font-family: 'Ubuntu', sans-serif; font-size: 11px;
    font-weight: 700; letter-spacing: 2px; text-transform: uppercase;
    cursor: pointer; margin-top: 12px;
    border-radius: 999px;
    box-shadow: 0 4px 14px rgba(245,13,116,0.30);
    transition: background 0.2s, box-shadow 0.2s;
  }
  .btn-login:hover { background: var(--raspberry-light); box-shadow: 0 6px 18px rgba(245,13,116,0.36); }
  .error {
    background: var(--red-soft);
    border: 1px solid rgba(218,41,41,0.30);
    color: var(--red);
    font-size: 12px; padding: 10px 14px; margin-bottom: 18px;
    border-radius: 8px;
    font-family: 'General Sans', sans-serif;
  }
  .footer-note {
    font-family: 'General Sans', sans-serif;
    font-size: 10px; color: var(--text-soft);
    margin-top: 28px; text-align: center; letter-spacing: 0.5px;
  }
  .site-footer {
    background: var(--navy-deep);
    color: rgba(245,247,250,0.72);
    padding: 18px 32px;
    text-align: center;
    z-index: 1; position: relative;
  }
  .site-footer .footer-brand {
    font-family: 'Ubuntu', sans-serif;
    font-size: 13px; font-weight: 700; color: #F5F7FA;
    margin-right: 12px;
  }
  .site-footer .footer-brand span { color: var(--raspberry-light); }
  .site-footer .footer-copy {
    font-family: 'General Sans', sans-serif;
    font-size: 10px; opacity: 0.85; letter-spacing: 0.5px;
  }
  @keyframes fadeUp { from { opacity: 0; transform: translateY(16px); } to { opacity: 1; transform: translateY(0); } }
</style>
</head>
<body>

<header class="site-header">
  <div class="header-inner">
    <div class="logo-mark" aria-label="Brightday logomark">
      <svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg" role="img" aria-hidden="true">
        <g transform="translate(50,50)">
          <g fill="#F50D74" fill-opacity="0.78">
            <rect x="-13" y="-34" width="26" height="26" rx="7" transform="rotate(0)"/>
            <rect x="-13" y="-34" width="26" height="26" rx="7" transform="rotate(45)"/>
            <rect x="-13" y="-34" width="26" height="26" rx="7" transform="rotate(90)"/>
            <rect x="-13" y="-34" width="26" height="26" rx="7" transform="rotate(135)"/>
            <rect x="-13" y="-34" width="26" height="26" rx="7" transform="rotate(180)"/>
            <rect x="-13" y="-34" width="26" height="26" rx="7" transform="rotate(225)"/>
            <rect x="-13" y="-34" width="26" height="26" rx="7" transform="rotate(270)"/>
            <rect x="-13" y="-34" width="26" height="26" rx="7" transform="rotate(315)"/>
          </g>
        </g>
      </svg>
    </div>
    <div class="header-text">
      <span class="brand">brightday</span>
      <span class="sub">SOA Completion Agent</span>
    </div>
  </div>
</header>

<main class="login-wrap">
  <div class="card">
    <h1>Sign <em>in</em></h1>
    <p class="subtitle">Internal access only. Enter your credentials to continue.</p>
    {% if error %}
    <div class="error">{{ error }}</div>
    {% endif %}
    <form method="POST" action="/login">
      <div class="field">
        <label>Username</label>
        <input type="text" name="username" autocomplete="username" autofocus required>
      </div>
      <div class="field">
        <label>Password</label>
        <input type="password" name="password" autocomplete="current-password" required>
      </div>
      <button class="btn-login" type="submit">Sign In →</button>
    </form>
    <p class="footer-note">Brightday Australia · ABN 45 674 252 905</p>
  </div>
</main>

<footer class="site-footer">
  <span class="footer-brand">brightday <span>Australia</span></span>
  <span class="footer-copy">ABN 45 674 252 905 · 260 Spencer Street, Melbourne · Internal Use Only</span>
</footer>

</body>
</html>"""


# ─────────────────────────────────────────────
# FLASK ROUTES
# ─────────────────────────────────────────────

@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        if check_password(username, password):
            session["authenticated"] = True
            session["username"] = username.lower()
            return redirect(url_for("tool"))
        else:
            error = "Invalid username or password."
    return render_template_string(LOGIN_HTML, error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
def tool():
    if not logged_in():
        return redirect(url_for("login"))
    html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "index.html")
    with open(html_path, "r") as f:
        return f.read()


@app.route("/process", methods=["POST"])
def process():
    if not logged_in():
        return jsonify({"error": "Not authenticated"}), 401
    try:
        if "fact_finder" not in request.files:
            return jsonify({"error": "Missing Fact Finder file"}), 400
        if "soa_template" not in request.files:
            return jsonify({"error": "Missing SOA Template file"}), 400

        risk_profile   = request.form.get("risk_profile", "").strip()
        scenario       = request.form.get("scenario", "").strip()
        no_insurance   = request.form.get("no_insurance", "false").lower() == "true"
        no_trauma      = request.form.get("no_trauma", "false").lower() == "true"
        no_salsac      = request.form.get("no_salsac", "false").lower() == "true"
        insurance_only = request.form.get("insurance_only", "false").lower() == "true"

        if not risk_profile:
            return jsonify({"error": "Risk profile must be selected"}), 400
        if not scenario:
            return jsonify({"error": "Scenario must be selected"}), 400
        if scenario not in {"1", "2", "3", "4", "5", "6"}:
            return jsonify({"error": "Scenario must be 1-6"}), 400

        ff_bytes       = request.files["fact_finder"].read()
        template_bytes = request.files["soa_template"].read()

        # Read fact finder then free its bytes
        data, conditionals = read_fact_finder(
            ff_bytes, risk_profile, no_insurance,
            no_trauma_flag=no_trauma,
            no_salsac_flag=no_salsac,
            insurance_only_flag=insurance_only,
            scenario=scenario,
        )
        del ff_bytes

        # Process SOA (template_bytes freed inside process_soa)
        out = process_soa(template_bytes, data, conditionals)

        client_name = data.get("{{ClientFullName}}", "Client")
        today = date.today().strftime("%Y%m%d")
        filename = f"SOA_Draft_{client_name.replace(' ','_')}_{today}.docx"

        return send_file(
            out,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except KeyError as e:
        return jsonify({"error": f"Fact Finder tab not found or unexpected structure: {e}"}), 400
    except Exception as e:
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host="0.0.0.0", port=port)
