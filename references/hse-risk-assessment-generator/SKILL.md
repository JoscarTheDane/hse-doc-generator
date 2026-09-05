---
name: hse-risk-assessment-generator
description: Generate a project-specific JSA/HIRA Excel workbook from SOWs, procedures, and work scopes. Reads source documents, extracts tasks/hazards, populates a 18-column risk assessment matrix with initial/residual risk scoring, RRMs, and action parties. Self-contained — all column logic, scoring, and formatting is in the Python script so it works on any Hermes setup.
---

# HSE Risk Assessment Generator (JSA/HIRA)

# HSE Risk Assessment Generator (JSA/HIRA)

Generates a formatted Excel JSA/HIRA workbook from procedure documents and SOWs. Reads source documents, extracts tasks/hazards, populates a risk assessment matrix with initial/residual risk scoring, RRMs, and action parties. Self-contained — all column logic, scoring, and formatting is in the Python script so it works on any Hermes setup.

## When to Use

When given a project SOW (PDF) and one or more procedure documents (DOCX) and asked to produce a JSA or HIRA. The user says "make me a JSA" or "generate a risk assessment."

## Default Output: 12-Column HIRA (Atlantic Marine / Marine Platforms Style)

This is the user's primary and expected format. Produces a **3-sheet Excel workbook**:

### Sheet 1: `HIRA 1 {WORKSITE_SHORT}` — Main Risk Assessment

**12-column layout:**

| Col | Width | Header (row 12) | Notes |
|-----|-------|----------------|-------|
| A | 7 | No. | Decimal task numbering (1.0, 2.1, 2.2, 5.3...) |
| B | 32 | Task / Activity | Full task description |
| C | 34 | HAZARD (source of harm) | Bullet list of conditions/sources |
| D | 30 | Threats / Scenario (how it could happen) | How each hazard could be triggered |
| E | 34 | Top Event & CONSEQUENCE (outcome) | Top event → consequences |
| F | 5 | P | Probability score (compound format) |
| G | 13 | S | Severity score (compound format) |
| H | 13 | RP | Residual Probability |
| I | 13 | RS | Residual Severity |
| J | 48 | Existing Controls / Risk Reduction Measures | Numbered list |
| K | 30 | Recovery Measures | What to do if controls fail |
| L | 24 | Action Party | Responsible parties |

**Header section (rows 1-11):**
- Row 1: Title merged across A-L: "HAZARD IDENTIFICATION & RISK ASSESSMENT (HIRA) — LEVEL 1" then project name
- Row 2: blank spacer
- Row 3: "Client:" | Client name
- Row 4: "Contractor:" | Contractor name
- Row 5: "Vessel:" | Vessel name
- Row 6: "Location:" | Location
- Row 7: "Worksite:" | Worksite description
- Row 8: "Document Ref:" | Doc ref (AM-{client_abbr}-{vessel_abbr}-HIRA-{year})
- Row 9: "Date:" | Date (DD MM YYYY)
- Rows 10-11: blank spacers
- Row 12: Column headers — dark blue fill (#2F5496), white bold text

**Phase header rows:**
- Format: "Phase N: Phase Name"
- Blue fill (#4472C4), bold text, spans all columns A-L
- Separates work phases
- Tasks under a phase share the same parent number (phase 5 → tasks 5.1, 5.2, 5.3...)

**Row 13 onwards:** Data rows with thin borders, wrap-text, vertical top alignment.

### Sheet 2: `RAM` — Risk Assessment Matrix

**5×A-E custom matrix** (NOT standard 1-25 multiplication):

- Left column: Severity levels 0-5 with consequence descriptions (People, Assets, Community, Environment)
- Top row: Likelihood columns A-E
  - A = Never heard of in the Industry
  - B = Heard of in the Industry
  - C = Has happened in the Organization or >1/yr in the Industry
  - D = Has happened at the Location or >1/yr in the Organization
  - E = Has happened >1/yr at the Location

**Severity levels:**
- 0 = No injury or health effect / No damage
- 1 = Slight injury or health effect / Slight damage — NTC or First Aid case
- 2 = Minor injury or health effect / Minor damage — Medical Treatment Case
- 3 = Major injury or health effect / Moderate damage — Lost Workday Case or Restricted Work Case
- 4 = PTD or up to 3 fatalities / Major damage
- 5 = More than 3 fatalities / Massive damage

**Colour-coded risk bands** (applied to the RAM matrix and to P/S cells in HIRA):
- Low risk cells → Yellow (#FFFFFF00)
- Medium risk → Yellow (#FFFFFF00)
- High risk → Red (#FFFF0000)

### Sheet 3: `Attendance` — Sign-in Sheet

- Header: "HIRA LEVEL 1 — ATTENDANCE LIST"
- Project name, vessel name, date
- Table: S/N | Name | Role | Signature
- Pre-populated with project team roles (to be filled in by attendees)

## P and S Compound Format (CRITICAL — do not use plain 1-5)

P and S values use a **compound format** like `2B`, `3C`, `1B` — combining a number and letter:

- **Number** = severity level (1-5 from the RAM matrix rows)
- **Letter** = likelihood band (A-E from the RAM matrix columns)

Examples from real data:
- `2B` = severity 2 (Minor injury), likelihood B (Heard of in Industry)
- `3C` = severity 3 (Major injury), likelihood C (Happened in Org)
- `3E` = severity 3 (Major injury), likelihood E (Happened >1/yr at Location)
- `1B` = severity 1 (Slight injury), likelihood B (Heard of in Industry)
- `2C` = severity 2 (Minor injury), likelihood C (Happened in Org)

In the example file, **RP and RS are all 0** — the user's format does not calculate separate residual scores. The P and S compound codes represent the *current assessed* risk position on the matrix grid.

## How It Works

The generator reads procedure documents (DOCX) and SOW (PDF), identifies discrete work phases and tasks from document structure (paragraphs, tables, headings), then constructs a HIRA with:

- **12-column HIRA matrix** (see format above)
- **Custom A-E × 0-5 risk matrix** (not standard 1-25 multiplication)
- **3 sheets**: HIRA main, RAM reference, Attendance
- **Work phases** extracted from document structure

## Variables (set per project)

| Variable | Example |
|---|---|
| CLIENT | Marine Platforms Limited |
| CONTRACTOR | Atlantic Marine & Oilfield Services Ltd |
| PROJECT_NAME | Single Anchor Loading Flange Disconnection and Reconnection |
| VESSEL | AFRICAN CONCEPT |
| LOCATION | AKPO Field / OML 130, Offshore Nigeria |
| WORKSITE | SAL (Single Anchor Loading) Buoy |
| DOCUMENT_REF | AM-MP-AC-HIRA-2026 |
| DATE | 05 07 2026 |

## ⚠️ CRITICAL: Hazard → Threat → Consequence Distinction

Apply this distinction rigorously for EVERY task row. This is the single most common error.

- **HAZARD** (col C) = the *source* or condition with potential to cause harm. A noun or noun phrase. Answers: *what is the source of harm?*
  - Examples: "Suspended load", "Live electrical systems (electrocution risk)", "DP loss of position", "Umbilical entanglement", "Strong underwater current (>0.7kts)", "Sharp edges / marine growth", "Defective tools (cracked flogging spanner)"

- **THREAT / SCENARIO** (col D) = *how* the hazard could be triggered or materialise. Answers: *how could this happen?*
  - Examples: "HMPE sling damaged or overloaded — no SWL label", "DP software failure / PRS dropout", "Disturbed silt — visibility reduced to zero", "Diver working in confined space between riser and buoy", "Shackle pin not secured — moused incorrectly", "Bolts corroded and seized — excessive force required"

- **TOP EVENT & CONSEQUENCE** (col E) = the *outcome* if the hazard materialises. Always describes harm, loss, or delay. Includes a "→" arrow linking top event to consequences.
  - Format: Top event description on first line → consequences listed with bullet points
  - Examples: "Diver in water with compromised station-keeping / deployment equipment failure → • Diver fatality / serious injury • DCI requiring hyperbaric treatment"

**Self-check for every row:** Read the hazard cell. Does it describe a source/condition, or an outcome? If it sounds like an outcome, it belongs in consequences. If the hazards column starts reading like a list of bad things that happen rather than a list of things that ARE present, you have swapped them.

**Common errors:**
- ❌ "Personnel injury, equipment damage" as a hazard → these are consequences
- ❌ "Schedule impact" as a hazard → it's a consequence of weather or equipment failure
- ❌ Bundling "DP loss of position, LARS winch failure, diver DCI" as one hazard → split: "DP loss of position" and "LARS winch failure" are hazards; "DCI" is a consequence
- ❌ Writing hazards as full sentences ("The diver may become entangled") → hazards are conditions/sources; use bullet lists
- ❌ Missing the "→" arrow in consequence column that separates top event from outcome

## Content Style Guidelines

### Hazards (Col C)
- Bullet-pointed list of conditions/sources
- Each hazard is a short noun phrase, NOT a sentence
- Include specific details where relevant (e.g., "Defective tools (cracked flogging spanner, damaged hammer)")

### Threats / Scenario (Col D)
- Describes the mechanism of failure
- Include contributing factors (e.g., "Poor visibility — tactile work only", "Diver exceeding safe bottom time per dive tables")
- Can be a short paragraph or bullet list

### Top Event & Consequence (Col E)
- Top event: what goes wrong (e.g., "Diver in water with compromised station-keeping")
- "→" arrow separator
- Consequences: bullet-listed outcomes with severity detail
- Example: "Dropped load / uncontrolled swing during lift → • Personnel injury (crush, impact) / fatality • Damage to dive equipment (DDC, LARS)"

### Existing Controls / RRM (Col J)
- Numbered list (1., 2., 3., ...)
- Include IMCA references where applicable (D010, D078, D067, LR006, M117, M190, etc.)
- Be specific and actionable, not generic
- Examples from the example file show 7-12 controls per task

### Recovery Measures (Col K)
- Concise, imperative statements
- Examples: "Emergency basket recovery via LARS. Standby diver deployed if required. DMT assessment. DDC prepped for immediate use. ERP Tier 2 — medevac activated if serious."
- List the immediate actions, then escalation path

### Action Party (Col L)
- Slash-separated list of roles
- Examples: "Diving Supervisor / Tender / Diver / Standby Diver"
- "PIC / Deck Crew / Crane Operator / Diving Supervisor / OCM"

## Work Phase Structure

Organise the HIRA into logical phases that match the project workflow:

1. **Pre-mobilisation & Personnel** — competency verification, cert checks, medical fitness
2. **Vessel Mobilisation & Integration** — equipment transfer, dive spread setup, system testing
3. **Transit to Field** — sailing, weather, security
4. **DP Setup & Positioning** — beacon deployment, PRS test, ASOG, weather windows
5. **Preparation to Dive** — pre-dive checks, LARS test, gas verification, comms, dress-in
6. **Diver Intervention (main task)** — deployment, excursion, task execution, recovery
7. **Post-Dive Close-out** — gas recharge, equipment inspection, log completion

Each phase gets a blue header row and its tasks use decimal numbering (5.1, 5.2, 5.3...).

## Hazard Templates (per work type)

### Lifting Operations
- Suspended load
- Uncertified/defective lifting equipment
- Congested deck / limited workspace
- Manual handling of heavy equipment
- Poor communication between crane operator and deck crew
- Inadequate lift plan / inexperienced banksman

### DP Diving Operations (LARS / Surface Supplied)

When divers deploy from a DP vessel via LARS, apply IMCA D078 (umbilical management) and IMCA D010 (DP diving) controls. Key hazards:

- DP loss of position (drive-off / drift-off)
- LARS winch wire / A-frame failure
- Umbilical entanglement during descent/recovery
- Vessel thruster proximity to diver
- Diver DCI from uncontrolled descent rate
- Weather deterioration mid-dive
- Basket snag on vessel hull
- Umbilical snag on subsea structure
- Strong underwater current (>0.7kts)
- Poor visibility (zero-vis conditions)

**IMCA D078 controls to include for DP diving tasks:**
- Working diver umbilical physically restrained — cannot reach within 5m of any hazard (thruster, propeller, sea chest)
- Standby diver umbilical restrained at 3m from hazard, 2m beyond working diver (D078 principles 4 & 5)
- D-ring demarcation system per approved vessel excursion plan
- Umbilical marked at 5m/10m intervals (red/black tape system per D078 §8.1)
- Vessel hazard diagram on bridge and in dive control per D078 §10
- 3m karabiner maintaining diver proximity to swim line; all karabiners lockable type (D078 §8.2)
- Weak link used for diver attachment to subsea structure per IMCA D058 / D078
- Active umbilical tending — tender monitors tension and payout (D078 §5)

### Diving Operations (General)
- Faulty dive equipment / incorrect setup
- DCI (Decompression Illness)
- Uncontrolled ascent/descent
- Umbilical entanglement
- Delta P (differential pressure)
- Marine life / environmental hazards
- Co-activity / SIMOPS

### Contaminated Air Supply

### Subsea Task Execution (flange work, bolting, etc.)
- Hand tool injury (pinch points, impact)
- Defective tools (cracked flogging spanner, damaged hammer)
- Dropped objects (bolts, spanner, hammer)
- Diver fatigue / extended bottom time
- Sharp edges on flange / debris
- Restricted work position — limited mobility

## IMCA References to Cite

Frequently used references in HIRA controls:
- IMCA D010 — Diving from DP vessels
- IMCA D018 — Examination and testing of diving plant
- IMCA D022/D023/D040 — DESIGN for air diving systems
- IMCA D043 — Gas cylinder marking
- IMCA D058 — Weak link attachment to subsea structure
- IMCA D067 — Underwater currents and diver safety
- IMCA D078 — Umbilical management (5m/3m rule, D-rings, markings, hazard diagrams)
- IMCA D084 — Diving contractor membership assessment
- IMCA LR006 — Lifting guidelines
- IMCA M117 — DP personnel training
- IMCA M190 — DP annual trials
- IMCA M273 — DP drills
- IOGP 411 — Diving recommended practice
- USN Rev7 — Dive tables
- BS EN 12012 — Air purity standard
- DMAC 15 — Medical equipment at dive sites

## Python Script (3-Sheet Excel Generator)

```python
#!/usr/bin/env python3
"""
HIRA Generator — 12-column Atlantic Marine / Marine Platforms format.
Produces 3 sheets: HIRA main matrix, RAM reference, Attendance list.
"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, numbers
from openpyxl.utils import get_column_letter

# === Style Definitions ===
HEADER_FONT = Font(bold=True, size=10, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
PHASE_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
PHASE_FONT = Font(bold=True, size=10, color="FFFFFF")
HEADER_FILL_INFO = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)
BODY_ALIGN = Alignment(wrap_text=True, vertical='top', horizontal='left')
CENTER_ALIGN = Alignment(wrap_text=True, vertical='center', horizontal='center')

# === Matrix Lookup ===
LIKELIHOOD_LABELS = {
    'A': 'Never heard of in the Industry',
    'B': 'Heard of in the Industry',
    'C': 'Has happened in the Organization or more than once per year in the Industry',
    'D': 'Has happened at the Location or more than once per year in the Organization',
    'E': 'Has happened more than once per year at the Location',
}

SEVERITY_DATA = {
    0: {'people': 'No injury or health effect', 'assets': 'No damage', 'community': 'No effect', 'env': 'No effect'},
    1: {'people': 'Slight injury or health effect', 'assets': 'Slight damage', 'community': 'Slight effect', 'env': 'Slight effect'},
    2: {'people': 'Minor injury or health effect', 'assets': 'Minor damage', 'community': 'Minor effect', 'env': 'Minor effect'},
    3: {'people': 'Major injury or health effect', 'assets': 'Moderate damage', 'community': 'Moderate effect', 'env': 'Moderate effect'},
    4: {'people': 'PTD or up to 3 fatalities', 'assets': 'Major damage', 'community': 'Major effect', 'env': 'Major effect'},
    5: {'people': 'More than 3 fatalities', 'assets': 'Massive damage', 'community': 'Massive effect', 'env': 'Massive effect'},
}

SEVERITY_DETAIL = {
    0: {'people': 'No Treatment Case or First Aid case', 'assets': 'Cost to rectify -Nil', 'env': 'No impact to the environment'},
    1: {'people': 'No Treatment Case or First Aid case', 'assets': 'Costs less than US $100,000', 'env': 'Slight environmental damage contained within the premises'},
    2: {'people': 'Medical Treatment Case', 'assets': 'Costs between US $100,000 and US $1 million', 'env': 'Minor environmental damage but no lasting effect'},
    3: {'people': 'Lost Workday Case or Restricted Work Case', 'assets': 'Costs between US $1 million and US $10 million', 'env': 'Limited environmental damage that will persist or require clean up'},
    4: {'people': 'Illness with irreversible health effect', 'assets': 'Costs between US $10 million and US $100 million', 'env': 'Severe environmental damage that will require extensive measures'},
    5: {'people': 'Cancer in a large exposed population', 'assets': 'Costs in excess of US $100 million', 'env': 'Persistent severe environmental damage leading to loss of natural resources'},
}


def make_compound_score(severity_num, likelihood_letter):
    """Create compound P/S score like '2B', '3C', '1B'."""
    return f"{severity_num}{likelihood_letter}"


def create_hira_workbook(client, contractor, project, vessel, location, worksite, doc_ref, date, phases):
    """Create 3-sheet HIRA Excel workbook."""
    wb = openpyxl.Workbook()

    # === SHEET 1: HIRA Main Matrix ===
    ws = wb.active
    ws.title = "HIRA 1"

    # Column widths
    col_widths = {
        'A': 7.0, 'B': 32.0, 'C': 34.0, 'D': 30.0, 'E': 34.0,
        'F': 5.0, 'G': 13.0, 'H': 13.0, 'I': 13.0,
        'J': 48.0, 'K': 30.0, 'L': 24.0
    }
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # Row 1: Title
    ws.merge_cells('A1:L1')
    title = f"HAZARD IDENTIFICATION & RISK ASSESSMENT (HIRA) — LEVEL 1\n{project}"
    cell = ws['A1']
    cell.value = title
    cell.font = Font(bold=True, size=12)
    cell.alignment = Alignment(wrap_text=True, horizontal='left', vertical='center')
    ws.row_dimensions[1].height = 30

    # Row 2: spacer
    ws.row_dimensions[2].height = 6

    # Info rows (3-9)
    info_items = [
        ('Client:', client),
        ('Contractor:', contractor),
        ('Vessel:', vessel),
        ('Location:', location),
        ('Worksite:', worksite),
        ('Document Ref:', doc_ref),
        ('Date:', date),
    ]
    for i, (label, value) in enumerate(info_items):
        row = i + 3
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        ws.cell(row=row, column=1).fill = HEADER_FILL_INFO
        ws.cell(row=row, column=2, value=value)
        ws.merge_cells(f'B{row}:L{row}')
        ws.row_dimensions[row].height = 18

    # Spacer rows
    ws.row_dimensions[10].height = 6
    ws.row_dimensions[11].height = 6

    # Row 12: Column headers
    headers = [
        'No.', 'Task / Activity',
        'HAZARD\n(source of harm)', 'Threats / Scenario\n(how it could happen)',
        'Top Event & CONSEQUENCE\n(outcome)',
        'P', 'S', 'RP', 'RS',
        'Existing Controls /\nRisk Reduction Measures',
        'Recovery Measures', 'Action Party'
    ]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=12, column=col, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
        cell.border = THIN_BORDER

    # Write phases and tasks
    data_row = 13
    for phase in phases:
        # Phase header row
        ws.merge_cells(f'A{data_row}:L{data_row}')
        phase_cell = ws.cell(row=data_row, column=1, value=phase['name'])
        phase_cell.font = PHASE_FONT
        phase_cell.fill = PHASE_FILL
        phase_cell.alignment = Alignment(horizontal='left', vertical='center')
        for c in range(1, 13):
            ws.cell(row=data_row, column=c).border = THIN_BORDER
            ws.cell(row=data_row, column=c).fill = PHASE_FILL
        ws.row_dimensions[data_row].height = 20
        data_row += 1

        # Task rows
        for task in phase.get('tasks', []):
            ws.cell(row=data_row, column=1, value=task['no']).alignment = CENTER_ALIGN
            ws.cell(row=data_row, column=1).border = THIN_BORDER

            ws.cell(row=data_row, column=2, value=task['task']).alignment = BODY_ALIGN
            ws.cell(row=data_row, column=2).border = THIN_BORDER

            ws.cell(row=data_row, column=3, value=task['hazard']).alignment = BODY_ALIGN
            ws.cell(row=data_row, column=3).border = THIN_BORDER

            ws.cell(row=data_row, column=4, value=task['threat']).alignment = BODY_ALIGN
            ws.cell(row=data_row, column=4).border = THIN_BORDER

            # Top Event & Consequence — merge E column (but it's already single col in 12-col format)
            ws.cell(row=data_row, column=5, value=task['consequence']).alignment = BODY_ALIGN
            ws.cell(row=data_row, column=5).border = THIN_BORDER

            # P, S, RP, RS
            for col_idx, val in enumerate([task['p'], task['s'], task['rp'], task['rs']], 6):
                cell = ws.cell(row=data_row, column=col_idx, value=val)
                cell.alignment = CENTER_ALIGN
                cell.border = THIN_BORDER

            ws.cell(row=data_row, column=10, value=task['controls']).alignment = BODY_ALIGN
            ws.cell(row=data_row, column=10).border = THIN_BORDER

            ws.cell(row=data_row, column=11, value=task['recovery']).alignment = BODY_ALIGN
            ws.cell(row=data_row, column=11).border = THIN_BORDER

            ws.cell(row=data_row, column=12, value=task['action']).alignment = BODY_ALIGN
            ws.cell(row=data_row, column=12).border = THIN_BORDER

            # Row height
            max_lines = max(
                str(task.get('hazard', '')).count('\n') + 1,
                str(task.get('threat', '')).count('\n') + 1,
                str(task.get('consequence', '')).count('\n') + 1,
                str(task.get('controls', '')).count('\n') + 1,
                3
            )
            ws.row_dimensions[data_row].height = max(45, max_lines * 12)
            data_row += 1

    # === SHEET 2: RAM ===
    ws_ram = wb.create_sheet("RAM")
    build_ram_sheet(ws_ram)

    # === SHEET 3: Attendance ===
    ws_att = wb.create_sheet("Attendance")
    build_attendance_sheet(ws_att, project, vessel, date)

    return wb


def build_ram_sheet(ws):
    """Build the RAM (Risk Assessment Matrix) reference sheet."""
    # Title row
    ws.merge_cells('A1:F1')
    ws['A1'] = 'SEVERITY'
    ws['A1'].fill = HEADER_FILL_INFO
    ws['A1'].font = Font(bold=True)

    ws.merge_cells('F1:J1')
    ws['F1'] = 'INCREASING LIKELIHOOD'
    ws['F1'].fill = HEADER_FILL_INFO
    ws['F1'].font = Font(bold=True)

    # Column headers
    consequence_headers = ['', 'CONSEQUENCES', '', '', '']
    for c, h in enumerate(consequence_headers, 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font = Font(bold=True)

    ws.cell(row=2, column=2, value='People')
    ws.cell(row=2, column=3, value='Assets')
    ws.cell(row=2, column=4, value='Community')
    ws.cell(row=2, column=5, value='Environment')

    for c, letter in enumerate(['A', 'B', 'C', 'D', 'E'], 6):
        ws.cell(row=2, column=c, value=letter)

    # Likelihood descriptions (row 3)
    for c, (letter, desc) in enumerate(LIKELIHOOD_LABELS.items(), 6):
        ws.cell(row=3, column=c, value=desc)

    # Severity rows (0-5)
    for sev_idx, sev_level in enumerate(range(0, 6)):
        data_row = 4 + sev_idx * 2  # rows 4,6,8,10,12,14
        detail_row = 5 + sev_idx * 2  # rows 5,7,9,11,13,15

        # Main data row
        ws.cell(row=data_row, column=1, value=sev_level)
        sd = SEVERITY_DATA[sev_level]
        ws.cell(row=data_row, column=2, value=sd['people'])
        ws.cell(row=data_row, column=3, value=sd['assets'])
        ws.cell(row=data_row, column=4, value=sd['community'])
        ws.cell(row=data_row, column=5, value=sd['env'])

        # Matrix cells
        for c, letter in enumerate(['A', 'B', 'C', 'D', 'E'], 6):
            cell = ws.cell(row=data_row, column=c, value=f'{letter}{sev_level}')

        # Detail row (second row per severity)
        sd2 = SEVERITY_DETAIL.get(sev_level, {})
        ws.cell(row=detail_row, column=2, value=sd2.get('people', ''))
        ws.cell(row=detail_row, column=3, value=sd2.get('assets', ''))
        ws.cell(row=detail_row, column=5, value=sd2.get('env', ''))

    # Row heights
    ws.row_dimensions[4].height = 22
    ws.row_dimensions[6].height = 22
    ws.row_dimensions[8].height = 22
    ws.row_dimensions[10].height = 22
    ws.row_dimensions[12].height = 22
    ws.row_dimensions[14].height = 22

    # Column widths
    ws.column_dimensions['A'].width = 7
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 40
    ws.column_dimensions['D'].width = 40
    ws.column_dimensions['E'].width = 40


def build_attendance_sheet(ws, project, vessel, date):
    """Build the Attendance sign-in sheet."""
    ws.merge_cells('A1:C1')
    ws['A1'] = 'HIRA LEVEL 1 — ATTENDANCE LIST'
    ws['A1'].font = Font(bold=True, size=12)

    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 6

    # Project line
    ws.row_dimensions[3].height = 22
    ws.merge_cells('A3:C3')
    ws['A3'] = f"Project: {project}   |   Vessel: {vessel}   |   {date}"

    ws.row_dimensions[4].height = 6

    # Table header
    headers = ['S/N', 'Name', 'Role', 'Signature']
    widths = [6, 35, 45, 20]
    for c, (h, w) in enumerate(zip(headers, widths), 1):
        cell = ws.cell(row=5, column=c, value=h)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL_INFO
        cell.border = THIN_BORDER
        ws.column_dimensions[get_column_letter(c)].width = w

    # Pre-populated roles
    roles = [
        'Project Coordinator', 'Project Manager', 'Executive Director (Operations)',
        'Dive Supervisor', 'HSE Lead', 'Equipment Manager', 'Onshore Manager',
        'Hyperbaric Doctor', 'Diving CSR / Client Rep', 'Vessel Master',
        'Marine Superintendent', 'DP Operator', 'Crane Operator',
        'Client Project Manager', 'Client HSE',
    ]
    for i, role in enumerate(roles, 1):
        row = i + 5
        ws.cell(row=row, column=1, value=i).alignment = CENTER_ALIGN
        ws.cell(row=row, column=1).border = THIN_BORDER
        ws.cell(row=row, column=2).border = THIN_BORDER
        ws.cell(row=row, column=3, value=role).border = THIN_BORDER
        ws.cell(row=row, column=4).border = THIN_BORDER
        ws.row_dimensions[row].height = 18


# ===== HAZARD LIBRARY =====
# Reusable hazard profiles for common subsea tasks

HAZARD_DP_DIVING = {
    'hazards': (
        "• DP loss of position (drive-off / drift-off)\n"
        "• LARS winch wire / A-frame failure\n"
        "• Umbilical entanglement during descent\n"
        "• Diver DCI from uncontrolled descent rate\n"
        "• Vessel thruster proximity to diver\n"
        "• Weather deterioration mid-dive"
    ),
    'threat': (
        "DP software failure / PRS dropout.\n"
        "LARS wire parted — shock load or fatigue.\n"
        "Umbilical trailing line caught on vessel hull.\n"
        "Basket descent too fast — winch brake failure.\n"
        "Squall line passing through field."
    ),
    'consequence': (
        "Diver in water with compromised station-keeping / deployment equipment failure →\n"
        "• Diver fatality / serious injury\n"
        "• DCI requiring hyperbaric treatment\n"
        "• Umbilical severed — loss of gas supply to diver\n"
        "• Vessel collision with SAL buoy / floating hose\n"
        "• Full ERP activation — critical incident"
    ),
    'controls': (
        "1. DP reliability programme — trained DP operators, ASOG, FMEA, annual DP trials per IMCA M190.\n"
        "2. IMCA D010 diving from DP vessels — DP status comms protocol to Dive Supervisor.\n"
        "3. IMCA D078 — umbilical management (5m restraint from hazards, D-ring demarcation, 5m interval marking).\n"
        "4. Diver tracked via Mesotech survey beacon and CCTV hat camera.\n"
        "5. LARS PMS and D023 certification current.\n"
        "6. Active umbilical tending — tender monitors tension and payout (IMCA D078 §5).\n"
        "7. IMCA M273 DP drills conducted prior to diving ops.\n"
        "8. Standby diver dressed and in immediate readiness."
    ),
    'recovery': (
        "Emergency basket recovery via LARS.\n"
        "Standby diver deployed if required.\n"
        "DMT assessment. DDC prepped for immediate use.\n"
        "ERP Tier 2 — medevac activated if serious."
    ),
    'action': 'Diving Supervisor / LARS Operator / Tender / DMT / Vessel Master / DP Operator',
    'p': '3C', 's': '2C', 'rp': '0', 'rs': '0',
}

HAZARD_LIFTING = {
    'hazards': (
        "• Suspended load\n"
        "• Uncertified/defective lifting equipment\n"
        "• Congested deck / limited workspace\n"
        "• Manual handling of heavy equipment\n"
        "• Poor communication between crane operator and deck crew"
    ),
    'threat': (
        "Lifting gear certificate expired or colour code out of date.\n"
        "Inexperienced banksman / rigger.\n"
        "No approved lift plan.\n"
        "Vessel movement in port."
    ),
    'consequence': (
        "Dropped load / uncontrolled swing during lift →\n"
        "• Personnel injury (crush, impact) / fatality\n"
        "• Damage to dive equipment (DDC, LARS)\n"
        "• Damage to vessel deck\n"
        "• Project delay"
    ),
    'controls': (
        "1. Only certified lifting equipment with valid colour code used.\n"
        "2. Approved lift plan reviewed by PIC before each lift.\n"
        "3. Exclusion zones enforced under all suspended loads.\n"
        "4. Competent banksman and dedicated comms channel.\n"
        "5. Manual handling — TILE assessment, two-person lift where needed.\n"
        "6. Good housekeeping — clear walkways and landing zones.\n"
        "7. Pre-use inspection of all rigging accessories."
    ),
    'recovery': (
        "Stop work. First aid by vessel medic.\n"
        "Follow AM ERP for medevac if serious injury.\n"
        "Securing and isolation of damaged equipment."
    ),
    'action': 'PIC / Deck Crew / Crane Operator / Diving Supervisor / OCM',
    'p': '3C', 's': '2C', 'rp': '0', 'rs': '0',
}
