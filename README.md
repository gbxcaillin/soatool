# SOA Completion Agent
### Brightday Australia — Internal Tool

---

## What This Does

Reads a completed Fact Finder (.xlsx) and populates a Statement of Advice
template (.docx), inserting all mapped field values in **red font** for
mandatory adviser review.

Unmapped codes (e.g. `{{SuperGoal}}`, `{{NeedsAnalysisLifeInsurance}}`) are
left as raw `{{code}}` text for manual completion.

---

## Files

```
soa_agent/
├── app.py            ← Python backend (Flask)
├── index.html        ← Web UI (served by Flask)
├── requirements.txt  ← Python dependencies
└── README.md
```

---

## Setup (Local)

**Requirements:** Python 3.9+

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run the server
python app.py

# 3. Open in browser
http://localhost:5000
```

---

## Deploying to a Web Server

Any standard Python hosting will work. Recommended options:

### Option A — Azure App Service (simplest for Microsoft ecosystem)
1. Create an App Service (Python 3.11, Linux)
2. Upload all files via VS Code Azure extension or GitHub Actions
3. Set startup command: `gunicorn app:app`
4. Add `gunicorn` to requirements.txt

### Option B — AWS Elastic Beanstalk
1. Zip the folder contents (not the folder itself)
2. Create a new Elastic Beanstalk Python environment
3. Upload the zip

### Option C — Any VPS (DigitalOcean, Linode, etc.)
```bash
pip install gunicorn
gunicorn --bind 0.0.0.0:80 app:app
```
Use nginx as a reverse proxy in front of gunicorn for production.

---

## What's Mapped (Auto-populated in red)

| Code | Source |
|------|--------|
| `{{Title}}` | FF row 10 |
| `{{ClientFullName}}` | FF rows 11+12+13 |
| `{{ClientFirstName}}` | FF row 11 |
| `{{ClientDOB}}` | FF row 15 |
| `{{ClientAddress}}` | FF rows 18–21 |
| `{{ClientPhone}}` | FF row 16 |
| `{{ClientEmail}}` | FF row 17 |
| `{{ClientOccupation}}` | FF row 28 |
| `{{ClientSalary}}` | FF row 32 |
| `{{fld_SuperContribution}}` | FF row 32 × row 34 |
| `{{fld_SalarySacrifice}}` | FF row 35 |
| `{{CurrentSuperFunds}}` | FF row 92 across fund cols |
| `{{SpouseDOB}}` | FF row 47 |
| `{{SpouseIncome}}` | FF row 49 |
| `{{SpouseBalance}}` | FF row 50 |
| `{{NoDependants}}` | Spouse + dependants count |
| `{{fld_CurrentSuperannuationBalance}}` | FF row 94 sum |
| `{{CurrentLifeInsurance}}` | FF row 102 across funds |
| `{{CurrentTPDInsurance}}` | FF row 103 across funds |
| `{{CurrentIncomeProtectionPerMonth}}` | FF row 104 across funds |
| `{{CurrentIncomeProtectionWaitingPeriod}}` | FF row 105 across funds |
| `{{CurrentIncomeProtectionBenefitPeriod}}` | FF row 106 across funds |
| `{{CurrentSuperPremiums}}` | FF row 107 across funds |
| `{{ValueOfPrimaryResidence}}` | FF row 73 |
| `{{DebtOnPrimaryResidence}}` | FF row 74 |
| `{{ValueOfInvestmentProperty}}` | FF row 76 |
| `{{DebtOnInvestmentProperty}}` | FF row 77 |
| `{{OtherAsset1Value}}` | FF row 79 |
| `{{PersonalLoan1Value}}` | FF row 81 |
| `{{TotalAssetValue}}` | Sum of all assets + super |
| `{{TotalLiabilityValue}}` | Sum of all debts |
| `{{RetirementAge}}` | FF row 9 (or row 8 if row 9 blank) |
| `{{CurrentBalance}}` | FF row 94 sum |
| `{{CurrentAge}}` | Calculated from FF row 15 DOB |
| `{{CurrentDate}}` | System date at generation time |
| `{{AnnualisedSalarySacrificeAmount}}` | FF row 35 |
| `{{BindingDeathNominee}}` | FF row 62 cols B/D/F/H/J |
| `{{CurrentRiskProfile}}` | UI dropdown selection |

---

## Conditional Block Logic

| Condition | Logic |
|-----------|-------|
| `DeleteIfAgeGreaterThan55` | Delete section if client age >= 55 |
| `DeleteIfAgeLessThan55` | Delete section if client age < 55 |
| `DeleteIfBalanceBelow500k` | Delete section if total super balance < $500,000 |
| `DeleteIfNoCurrentInsurance` | Delete if FF row 100 all No/blank |
| `DeleteIfNoInsuranceAtAll` | Delete if UI "No Insurance" checkbox ticked |
| `DeleteifNoCurrentUnderwrittenInsurance` | Delete if FF row 108 no "Medically Underwritten" |

---

## Unmapped Codes (Left as raw `{{code}}`)

These remain in the document for manual completion:

- `{{Date}}`, `{{OtherAsset1}}`, `{{OtherAsset2}}`, `{{OtherAsset2Value}}`
- `{{PersonalLoan2Value}}`, `{{****PersonalLoan1}}`, `{{****PersonalLoan2}}`
- `{{NeedsAnalysisLifeInsurance}}`, `{{NeedsAnalysisTPD}}`, `{{NeedsAnalysisIP}}`, `{{NeedsAnalysisTrauma}}`
- `{{Tbl_SalarySacrifice}}`, `{{tbl_CurrentSuperFundsRiskProfilePerformance}}`
- `{{SuperGoal}}`, `{{InsuranceGoal}}`, `{{SalarySacrificeGoal}}`, `{{EstatePlanningGoal}}`, `{{RetirementGoal}}`
- `{{CurrentInsuer}}`, `{{SalarySacrificeAmount}}`, `{{SalarySacrificeFrequency}}`, `{{NetTaxSavings}}`
- All scoped insurance conditional tags

---

## Adding More Mappings Later

Open `app.py` and find the `data = { ... }` dictionary in `read_fact_finder()`.
Add a new line:

```python
"{{YourNewCode}}": cell(ROW_NUMBER, COLUMN_NUMBER),
```

Then remove that code from the `UNMAPPED_CODES` set near the top of the file.

---

## KYC Note Integration (Future)

The architecture is ready for a KYC note processing step.
When ready, add a third file upload to `index.html` and a
`read_kyc_note()` function to `app.py` that merges its output
into the `data` dict before `process_soa()` is called.
