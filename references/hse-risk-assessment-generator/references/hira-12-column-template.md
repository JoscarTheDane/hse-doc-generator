# 12-Column HIRA Format — Reference

This is the definitive reference for the Atlantic Marine / Marine Platforms HIRA format. Match the example Excel at `~/HSE/2026-07-01-Atlantic-marine/Decent Example of whats expected.xlsx` exactly.

## 3-Sheet Excel Workbook

### Sheet 1: `HIRA 1 {WORKSITE}` — Main Matrix

**Column structure (12 cols, width in chars):**
```
A: 7    No. (decimal task no: 1.0, 2.1, 5.3...)
B: 32   Task / Activity
C: 34   HAZARD (source of harm)
D: 30   Threats / Scenario (how it could happen)
E: 34   Top Event & CONSEQUENCE (outcome)
F: 5    P (compound score: e.g. 3C, 2B)
G: 13   S (compound score: e.g. 2C, 1B)
H: 13   RP
I: 13   RS
J: 48   Existing Controls / Risk Reduction Measures
K: 30   Recovery Measures
L: 24   Action Party
```

**Header section (rows 1-12):**
- Row 1: Merged A1:L1 — "HAZARD IDENTIFICATION & RISK ASSESSMENT (HIRA) — LEVEL 1\n{project name}"
- Row 2: blank spacer
- Row 3: "Client:" | Client name (fill F2F2F2 on col A)
- Row 4: "Contractor:" | Contractor name
- Row 5: "Vessel:" | Vessel name (ALL CAPS)
- Row 6: "Location:" | Location
- Row 7: "Worksite:" | Worksite description
- Row 8: "Document Ref:" | AM-{client_abbr}-{vessel_abbr}-HIRA-{year}
- Row 9: "Date:" | DD MM YYYY
- Row 10-11: blank spacers
- Row 12: Column headers — dark blue (#2F5496) fill, white bold text, centered, wrap

**Phase header rows:**
- "Phase N: Phase Name" — merged A-L
- Blue (#4472C4) fill, bold white text
- Tasks under phase N use decimal numbering: N.1, N.2, N.3...

**Data rows (row 13+):**
- Thin borders on all cells
- Wrap text, vertical top alignment for text columns
- Center alignment for No., P, S, RP, RS columns
- Row height auto-scaled based on line count

### Sheet 2: `RAM` — Risk Assessment Matrix

Custom A-E × 0-5 matrix (NOT standard P×S=1-25):

**Structure:**
- Row 1: "SEVERITY" merged A1:F1 + "INCREASING LIKELIHOOD" merged F1:J1 (light grey fill F2F2F2)
- Row 2: "CONSEQUENCES" header with sub-headers: People, Assets, Community, Environment (cols B-E), Likelihood letters A-E (cols F-J)
- Row 3: Likelihood descriptions under each letter
- Rows 4-15: Severity levels 0-5, each taking 2 rows (data row + detail row)

**Severity definitions:**
- 0: No injury or health effect / No damage
- 1: Slight injury — No Treatment Case or First Aid
- 2: Minor injury — Medical Treatment Case
- 3: Major injury — Lost Workday or Restricted Work Case
- 4: PTD or up to 3 fatalities — Major damage
- 5: More than 3 fatalities — Massive damage

**Likelihood scale:**
- A = Never heard of in the Industry
- B = Heard of in the Industry
- C = Has happened in the Organization or >1/yr in Industry
- D = Has happened at the Location or >1/yr in Organization
- E = Has happened >1/yr at the Location

**Colour coding in matrix:**
- Yellow (#FFFFFF00) = medium risk cells
- Red (#FFFF0000) = high risk cells

### Sheet 3: `Attendance` — Sign-in Sheet

- Header: "HIRA LEVEL 1 — ATTENDANCE LIST"
- Row 3: "Project: {name} | Vessel: {name} | {date}"
- Table: S/N | Name (blank) | Role (pre-populated) | Signature (blank)
- 15 pre-populated roles

## P/S Compound Score Format

P and S use compound format: `{number}{letter}`

- Number = severity level from RAM (1-5)
- Letter = likelihood band from RAM (A-E)

Examples:
- `2B` = severity 2 (Minor), likelihood B (Heard in Industry)
- `3C` = severity 3 (Major), likelihood C (Happened in Org)
- `3E` = severity 3 (Major), likelihood E (Happened >1/yr at Location)
- `1B` = severity 1 (Slight), likelihood B (Heard in Industry)
- `2C` = severity 2 (Minor), likelihood C (Happened in Org)

**RP and RS are all 0** in the user's format — the P/S compound codes represent the current assessed risk position.

## Hazard → Threat → Consequence Format

**HAZARD (Col C):** Bullet-pointed list of sources/conditions. Short noun phrases.
**THREAT (Col D):** Mechanism of failure. Short paragraphs or bullet lists.
**TOP EVENT & CONSEQUENCE (Col E):** Top event → consequences. Arrow separator, then bullet-listed outcomes.

## Style Notes

- Controls: numbered list with IMCA references
- Recovery: concise imperative statements
- Action Party: slash-separated roles
- Content is highly detailed — 7-12 controls per task typical
