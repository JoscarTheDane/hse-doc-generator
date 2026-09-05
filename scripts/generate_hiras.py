# AgentMail key: set AGENTMAIL_API_KEY env var or place key file at $HSE_HOME/.api_key
#!/usr/bin/env python3
"""
Generate HIRA Excel workbooks for 4 projects.
Uses openpyxl to produce 3-sheet workbooks: HIRA main, RAM, Attendance.
"""

import os
import sys
import json
from datetime import datetime

sys.path.insert(0, '/tmp/hira_venv/lib/python3.11/site-packages')

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# === Style Constants ===
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

SAVE_DIR = "/home/joshua/HSE"
ATTACHMENT_DIR = "/home/joshua/.hermes/skills/productivity/hse-risk-assessment-generator/scripts/temp_attachments"

def create_hira_workbook(client, contractor, project, vessel, location, worksite, doc_ref, date, phases, title_short=None):
    """Create 3-sheet HIRA Excel workbook."""
    wb = Workbook()

    # === SHEET 1: HIRA Main Matrix ===
    ws = wb.active
    ws.title = f"HIRA 1 {title_short or worksite[:15]}"

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
    """Build the RAM reference sheet."""
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
    # Set values BEFORE merging (merge_cells overwrites value on top-left)
    ws.cell(row=1, column=1, value='SEVERITY')
    ws.cell(row=1, column=1).fill = HEADER_FILL_INFO
    ws.cell(row=1, column=1).font = Font(bold=True)

    ws.cell(row=1, column=6, value='INCREASING LIKELIHOOD')
    ws.cell(row=1, column=6).fill = HEADER_FILL_INFO
    ws.cell(row=1, column=6).font = Font(bold=True)

    # Now merge
    ws.merge_cells('A1:F1')
    ws.merge_cells('F1:J1')

    # Column headers
    ws.cell(row=2, column=2, value='CONSEQUENCES')
    ws.cell(row=2, column=2).font = Font(bold=True)
    ws.cell(row=2, column=3, value='People')
    ws.cell(row=2, column=3).font = Font(bold=True)
    ws.cell(row=2, column=4, value='Assets')
    ws.cell(row=2, column=4).font = Font(bold=True)
    ws.cell(row=2, column=5, value='Community')
    ws.cell(row=2, column=5).font = Font(bold=True)

    for c, letter in enumerate(['A', 'B', 'C', 'D', 'E'], 6):
        ws.cell(row=2, column=c, value=letter)
        ws.cell(row=2, column=c).font = Font(bold=True)

    for c, (letter, desc) in enumerate(LIKELIHOOD_LABELS.items(), 6):
        ws.cell(row=3, column=c, value=desc)

    for sev_idx, sev_level in enumerate(range(0, 6)):
        data_row = 4 + sev_idx * 2
        detail_row = 5 + sev_idx * 2

        ws.cell(row=data_row, column=1, value=sev_level)
        sd = SEVERITY_DATA[sev_level]
        ws.cell(row=data_row, column=2, value=sd['people'])
        ws.cell(row=data_row, column=3, value=sd['assets'])
        ws.cell(row=data_row, column=4, value=sd['community'])
        ws.cell(row=data_row, column=5, value=sd['env'])

        for c, letter in enumerate(['A', 'B', 'C', 'D', 'E'], 6):
            ws.cell(row=data_row, column=c, value=f'{letter}{sev_level}')

        sd2 = SEVERITY_DETAIL.get(sev_level, {})
        ws.cell(row=detail_row, column=2, value=sd2.get('people', ''))
        ws.cell(row=detail_row, column=3, value=sd2.get('assets', ''))
        ws.cell(row=detail_row, column=5, value=sd2.get('env', ''))

    ws.row_dimensions[4].height = 22
    ws.row_dimensions[6].height = 22
    ws.row_dimensions[8].height = 22
    ws.row_dimensions[10].height = 22
    ws.row_dimensions[12].height = 22
    ws.row_dimensions[14].height = 22

    ws.column_dimensions['A'].width = 7
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 40
    ws.column_dimensions['D'].width = 40
    ws.column_dimensions['E'].width = 40


def build_attendance_sheet(ws, project, vessel, date):
    """Build the Attendance sign-in sheet."""
    from openpyxl.utils import get_column_letter

    ws.merge_cells('A1:C1')
    ws['A1'] = 'HIRA LEVEL 1 — ATTENDANCE LIST'
    ws['A1'].font = Font(bold=True, size=12)
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 6

    ws.row_dimensions[3].height = 22
    ws.merge_cells('A3:C3')
    ws['A3'] = f"Project: {project}   |   Vessel: {vessel}   |   {date}"
    ws.row_dimensions[4].height = 6

    headers = ['S/N', 'Name', 'Role', 'Signature']
    widths = [6, 35, 45, 20]
    for c, (h, w) in enumerate(zip(headers, widths), 1):
        cell = ws.cell(row=5, column=c, value=h)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL_INFO
        cell.border = THIN_BORDER
        ws.column_dimensions[get_column_letter(c)].width = w

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


# ==============================================================================
# PROJECT 1: Snake Island Wrecks Removal
# ==============================================================================
def gen_project_1():
    """Novadeal EKO FZE - Snake Island Wrecks and Debris Removal."""
    project = "Snake Island Container Terminal Development - Wreck & Debris Removal"
    client = "Novadeal EKO FZE (DEME Group)"
    contractor = "ConsultingSubsea / Approved Diving Contractor"
    vessel = "Diving Support Vessel (DSV)"
    location = "Snake Island, Lagos, Nigeria"
    worksite = "Snake Island Container Terminal Development Dredging Area"
    doc_ref = "6356-CIN-RFQ-001 Rev. 1.0"
    date = "25 07 2026"

    phases = []

    # Phase 1: Pre-Mobilisation & Planning
    phases.append({
        'name': 'Phase 1: Pre-Mobilisation & Planning',
        'tasks': [
            {
                'no': '1.1',
                'task': 'Mobilisation of diving equipment and personnel to worksite',
                'hazard': '• Manual handling of heavy dive equipment\n• Uncertified/defective lifting equipment\n• Poor housekeeping on vessel deck',
                'threat': 'Equipment certificates expired or colour code out of date.\nInexperienced riggers.\nNo approved lift plan for deck operations.',
                'consequence': 'Dropped load / equipment damage →\n• Personnel injury (crush, impact)\n• Damage to critical dive equipment\n• Project delay due to equipment unavailability',
                'controls': '1. Only certified lifting equipment with valid colour code used.\n2. Approved lift plan reviewed before each lift.\n3. Exclusion zones enforced under suspended loads.\n4. Competent banksman and dedicated comms channel.\n5. Manual handling — TILE assessment, two-person lift where needed.\n6. Pre-use inspection of all rigging accessories.\n7. IMCA LR006 lifting guidelines compliance.',
                'recovery': 'Stop work. First aid by vessel medic.\nSecure and isolate damaged equipment.\nReplace with spare certified equipment.',
                'action': 'Deck Crew / Rigger / Crane Operator / Diving Supervisor / OCM',
                'p': '3C', 's': '2C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '1.2',
                'task': 'Pre-dive equipment inspection and commissioning',
                'hazard': '• Defective dive equipment\n• Contaminated air supply\n• Electrical equipment in wet environment',
                'threat': 'Compressor oil carryover or faulty filtration.\nDamaged umbilical from previous dive.\nFaulty electrical insulation on underwater equipment.',
                'consequence': 'Equipment failure during dive →\n• Loss of breathing gas to diver\n• Electric shock to diver\n• DCI requiring hyperbaric treatment\n• Loss of diver communication',
                'controls': '1. Compressor air analysis per BS EN 12012 — tested and certified before each dive.\n2. IMCA D018 examination and testing of diving plant — PMS current.\n3. Pre-dive check of all umbilicals — visual and pressure test.\n4. IMCA D043 gas cylinder marking and certification.\n5. IMCA D022/D023/D040 DESIGN audit current.\n6. All bail-out cylinders checked with calibrated Go/No-Go gauges.',
                'recovery': 'Terminate dive immediately.\nSwap to standby equipment.\nDMT assessment before resuming.\nERP — medevac if diver affected.',
                'action': 'Dive Technician / Diving Supervisor / Diver Medic / Tender',
                'p': '2C', 's': '3C', 'rp': '0', 'rs': '0',
            },
        ]
    })

    # Phase 2: Marine Survey & Area Clearance
    phases.append({
        'name': 'Phase 2: Marine Survey & Wreck Area Identification',
        'tasks': [
            {
                'no': '2.1',
                'task': 'MBES / SSS survey of wreck and debris fields',
                'hazard': '• ROV/diving SIMOPS\n• Vessel movement in confined area\n• Underwater obstacles and debris',
                'threat': 'ROV operating in same area as planned dive zone.\nSurvey vessel drifting in strong current.\nUncharted debris on seabed creating entanglement hazard.',
                'consequence': 'Collision with survey equipment or vessel →\n• Damage to survey equipment (MBES, SSS)\n• Umbilical entanglement with survey cable\n• Delay to dive operations\n• Environmental contamination from struck debris',
                'controls': '1. IMCA D054 — ROV/diving coordination procedure in place.\n2. Separate survey and diving windows where possible.\n3. Dedicated comms channel for SIMOPS coordination.\n4. Pre-survey area marked on chart with GPS coordinates.\n5. Current monitoring — operations suspended if >1 knot.\n6. Daily weather monitoring per IMCA D067.',
                'recovery': 'Suspend survey operations.\nRelocate vessel.\nReconnaissance dive to reassess area.\nEscalate to OIM/Master if risk cannot be controlled.',
                'action': 'Marine Supervisor / ROV Supervisor / Diving Supervisor / OCM',
                'p': '3C', 's': '2C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '2.2',
                'task': 'Magnetometry (MAG) survey for unexploded ordnance (UXO) assessment',
                'hazard': '• UXO exposure on seabed\n• Electromagnetic interference with dive equipment\n• Handling of hazardous materials',
                'threat': 'MAG anomaly indicating WWII-era ordnance.\nAcoustic survey equipment disturbing settled ordnance.\nAccidental contact with UXO during wreck removal.',
                'consequence': 'UXO explosion or detonation →\n• Diver fatality / mass casualties\n• Vessel damage\n• Total project shutdown\n• Regulatory penalties from NIMASA/NPA',
                'controls': '1. MAG survey completed by licensed UXO specialist before any diving.\n2. MAG results reviewed by qualified UXO officer.\n3. Exclusion zone established around any MAG anomalies.\n4. NIMASA notification — permit required before UXO handling.\n5. No diver deployment in UXO anomaly zones.\n6. Emergency response plan includes UXO incident scenario.\n7. IMCA guidance on submerged ordnance compliance.',
                'recovery': 'Evacuate all personnel from area.\nContact NIMASA and Explosive Ordnance Disposal (EOD).\nWait for EOD clearance before resuming.\nActivate ERP Tier 2.',
                'action': 'UXO Officer / Marine Supervisor / Diving Supervisor / OIM / NIMASA',
                'p': '2D', 's': '5C', 'rp': '0', 'rs': '0',
            },
        ]
    })

    # Phase 3: Wreck Assessment & Planning
    phases.append({
        'name': 'Phase 3: Wreck & Debris Assessment',
        'tasks': [
            {
                'no': '3.1',
                'task': 'Diver-assisted wreck survey and visual inspection',
                'hazard': '• Sharp edges on wreck debris\n• Entanglement in fishing net / cables\n• Confined spaces within wrecks\n• Marine life (stingers, urchins)\n• Poor visibility',
                'threat': 'Diver reaching into wreck structure.\nFishing gear entangling diver or umbilical.\nDiver entering partially collapsed wreck compartment.\nStrong bottom current (>0.7kts) per IMCA D067.\nSilt-out from previous operations reducing visibility to zero.',
                'consequence': 'Diver trapped / injured underwater →\n• Diver entrapment — requires standby diver rescue\n• Umbilical severed — loss of gas supply\n• Laceration / laceration injury\n• DCI from emergency ascent\n• Respiratory injury from disturbed sediments',
                'controls': '1. IMCA D010 — Diving from DP vessels protocol.\n2. IMCA D078 — umbilical management (5m restraint, D-ring demarcation).\n3. Diver tracked via USBL position tracking.\n4. Standby diver dressed and in immediate readiness.\n5. Pre-dive check of wreck approach route.\n6. Underwater lights for low visibility work.\n7. Marine life PPE — protective gloves and full coverage suit.\n8. IMCA D067 — current monitoring and diver safety.',
                'recovery': 'Emergency basket recovery via LARS.\nStandby diver deployed.\nDMT assessment.\nDDC prepped for 18-hour bend watch.\nERP — medevac if serious injury.',
                'action': 'Diving Supervisor / Tender / Diver / Standby Diver / Diver Medic',
                'p': '3C', 's': '3C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '3.2',
                'task': 'Wreck weight estimation and rigging plan development',
                'hazard': '• Uncertified wreck weight\n• Inadequate rigging design\n• Improper sling selection',
                'threat': 'Wreck weight significantly higher than estimated.\nSling capacity insufficient for actual load.\nNo formal lift plan for wreck recovery.',
                'consequence': 'Sling failure / uncontrolled drop →\n• Diver crushed by shifting wreck\n• Dropped load causing vessel damage\n• Umbilical damaged by falling wreck\n• Environmental contamination from rust/chemicals',
                'controls': '1. Wreck dimensions and density calculated by structural engineer.\n2. Lift plan prepared by competent person — approved before use.\n3. HMPE slings rated at minimum 2x estimated wreck weight.\n4. Choker hitch method for secure wrapping of wreck.\n5. IMCA LR006 — lifting guidelines compliance.\n6. Pre-lift meeting with all involved parties.\n7. Trial lift to 0.5m to verify stability.',
                'recovery': 'Stop all lifting operations.\nEvacuate diver from water.\nReassess wreck weight and rigging.\nEngage structural engineer for revised plan.\nActivate ERP if injury sustained.',
                'action': 'Diving Supervisor / Structural Engineer / PIC / Deck Crew / Diver',
                'p': '3C', 's': '3C', 'rp': '0', 'rs': '0',
            },
        ]
    })

    # Phase 4: Wreck Removal Operations
    phases.append({
        'name': 'Phase 4: Wreck & Debris Removal Operations',
        'tasks': [
            {
                'no': '4.1',
                'task': 'Diver-assisted rigging of wreck debris',
                'hazard': '• Suspended load\n• Sharp edges on wreck\n• Manual handling of slings\n• Underwater confined space\n• Silt-out / zero visibility',
                'threat': 'Wreck has sharp edges cutting HMPE slings.\nDiver working in confined space between wreck and seabed.\nSilt disturbed during sling placement reducing visibility.\nWreck unstable and shifts during rigging.',
                'consequence': 'Diver injured during rigging →\n• Crush injury from shifting wreck\n• Laceration from sharp edges\n• Dropped load / uncontrolled swing\n• Umbilical snagged on wreck\n• DCI from emergency response',
                'controls': '1. Edge protection (sleeving/wood) placed between wreck and slings.\n2. Two slings per wreck — balanced lift plan.\n3. IMCA D078 — umbilical tending and restraint.\n4. Active tender monitoring tension and payout.\n5. D-ring demarcation system for diver excursion.\n6. Pre-rigging diver briefing on hazard points.\n7. Standby diver in immediate readiness.\n8. Continuous comms between diver and supervisor.\n9. IMCA D058 — weak link used for diver attachment.',
                'recovery': 'Emergency basket recovery.\nFirst aid by vessel medic.\nAssess wreck stability before resuming.\nDMT review if DCI suspected.\nERP activation if serious injury.',
                'action': 'Diving Supervisor / Tender / Diver / Standby Diver / Deck Crew',
                'p': '3C', 's': '3C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '4.2',
                'task': 'Wreck recovery to surface using dive vessel crane/LARS',
                'hazard': '• Suspended load\n• Congested deck\n• Poor crane visibility\n• Weather deterioration\n• Vessel movement',
                'threat': 'Wreck swings uncontrollably due to current.\nCrane operator loses sight of diver signal.\nWeather deteriorates mid-recovery.\nWreck contacts vessel hull during recovery.',
                'consequence': 'Wreck impact with vessel or diver →\n• Vessel hull damage / flooding\n• Diver struck by swinging load\n• Umbilical severed by contact\n• Personnel injury on deck from load swing\n• Environmental contamination release',
                'controls': '1. Dedicated crane signalman with hard-wire comms to operator.\n2. Weather window monitored — crane ops suspended if wind >10 kts.\n3. Current limit 1 knot — suspended if exceeded (IMCA D067).\n4. Exclusion zone under suspended load.\n5. Tag line deployed from surface to control wreck swing.\n6. IMCA LR006 — approved lift plan for each wreck recovery.\n7. Pre-lift trial to 1m — verify stability and rigging.\n8. SWL clearly marked on all slings and shackles.\n9. Crane PMS and certification current.',
                'recovery': 'Stop crane operation.\nSecure wreck on deck with chocks and chains.\nInspect vessel for damage.\nReassess weather and current before resuming.\nFirst aid if injury sustained.',
                'action': 'Crane Operator / Signalman / Diving Supervisor / Marine Superintendent / OCM',
                'p': '3C', 's': '2C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '4.3',
                'task': 'Disposal of wreck debris to NPA Kirikiri dumpsite',
                'hazard': '• Transport of hazardous debris\n• NPA permit compliance\n• Marine pollution during transport',
                'threat': 'Debris containing oil/hydrocarbon residue released during transport.\nNPA permit not obtained or expired.\nDebris falling overboard during vessel transit.\nWreck not properly secured on deck.',
                'consequence': 'Environmental contamination →\n• Oil/hydrocarbon spill in Lagos waters\n• NPA fines and regulatory action\n• Project work stoppage\n• Environmental damage to marine ecosystem\n• Reputational damage to DEME/Novadeal',
                'controls': '1. NPA wreck disposal permit obtained before operations commence.\n2. Wreck debris washed and contained to prevent pollution.\n3. Hydrocarbon testing of debris before transport.\n4. Secure deck stowage — chains and lashing points verified.\n5. NPA Kirikiri dumpsite arrival documentation.\n6. IMCA D022/D023 — marine pollution prevention measures.\n7. Oil spill response equipment on standby during transport.\n8. Daily marine pollution monitoring report.',
                'recovery': 'Contain spill with booms.\nNotify NPA immediately.\nActivate marine pollution response plan.\nDeploy oil spill equipment.\nReport to environmental authority.',
                'action': 'Vessel Master / Environmental Officer / PIC / Diving Supervisor / OCM',
                'p': '2C', 's': '3D', 'rp': '0', 'rs': '0',
            },
            {
                'no': '4.4',
                'task': 'Post-removal survey and visual check',
                'hazard': '• ROV/diving SIMOPS\n• Underwater debris still present\n• Survey vessel in active area',
                'threat': 'Remaining debris not detected by initial survey.\nROV cable snagged on remaining wreck.\nSurvey vessel drifting into dive zone.',
                'consequence': 'Incomplete clearance →\n• Dredging equipment damage from missed debris\n• Survey equipment damage\n• Delay to dredging phase\n• Potential safety hazard for future operations',
                'controls': '1. Post-removal MBES survey — compare with pre-clearance survey.\n2. Visual inspection by diver or ROV of cleared area.\n3. Acceptance criteria per clause 2.1 of RFQ scope.\n4. Photo/video documentation of cleared seabed.\n5. Survey vessel kept clear of dive/ROV operations.\n6. IMCA D054 — SIMOPS coordination between survey and diving.',
                'recovery': 'Identify missed debris areas.\nDeploy diver/ROV for additional clearance.\nRe-survey area.\nEscalate to Novadeal Project Manager if scope exceeds estimate.',
                'action': 'Survey Supervisor / Diving Supervisor / ROV Supervisor / Project Manager',
                'p': '2C', 's': '2C', 'rp': '0', 'rs': '0',
            },
        ]
    })

    # Phase 5: Post-Dive Close-out
    phases.append({
        'name': 'Phase 5: Post-Dive Close-out & Demobilisation',
        'tasks': [
            {
                'no': '5.1',
                'task': 'Post-dive equipment inspection and gas recharge',
                'hazard': '• High pressure air systems\n• Hot compressor surfaces\n• Gas cylinder handling',
                'threat': 'Overpressurised cylinder — safety relief activated.\nCompressor hot surface causing burns.\nGas cylinder dropped during handling.\nContaminated air from faulty compressor.',
                'consequence': 'Equipment damage or injury →\n• Cylinder explosion or rupture\n• Burns from hot compressor\n• Crush injury from dropped cylinder\n• Contaminated air affecting next divers',
                'controls': '1. All gas cylinders pressure tested per IMCA D043.\n2. Compressor checked and serviced per PMS.\n3. Air analysis before next dive — O2, CO, CO2 levels.\n4. IMCA D018 — examination and testing of diving plant.\n5. BS EN 12012 — breathing air purity standard.\n6. Cylinder handling — two-person lift, proper trolley.\n7. Post-dive equipment checklist completed and signed.',
                'recovery': 'Isolate contaminated air system.\nFlush and retest compressor.\nReplace compromised cylinders.\nDMT assessment before resuming diving.\nNotify Diving Superintendent.',
                'action': 'Dive Technician / Diving Supervisor / Diver Medic / Deck Crew',
                'p': '2C', 's': '2C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '5.2',
                'task': 'Demobilisation of diving equipment and personnel',
                'hazard': '• Demobilisation lifting operations\n• Vessel departure in sea state\n• Equipment left on deck',
                'threat': 'Demob crane operations in port with vessel movement.\nEquipment not properly secured — falls overboard.\nSea state exceeds safe transit limit.',
                'consequence': 'Lost equipment / personnel injury →\n• Dive equipment lost overboard\n• Vessel damage during departure\n• Injury during demob operations\n• Project delay',
                'controls': '1. Demob lift plan prepared and approved.\n2. Equipment secured before vessel departure.\n3. Sea state check — transit suspended if swell >1.5m.\n4. All loose equipment stowed or lashed.\n5. Vessel Master approval for departure.\n6. IMCA LR006 — lifting during vessel transit.\n7. Post-dive equipment inventory check.',
                'recovery': 'Stop vessel departure.\nSecure equipment.\nReassess sea state.\nRecovery dive if equipment lost overboard.\nReport to Novadeal Project Manager.',
                'action': 'Vessel Master / Diving Supervisor / PIC / Deck Crew / Crane Operator',
                'p': '2C', 's': '2C', 'rp': '0', 'rs': '0',
            },
        ]
    })

    wb = create_hira_workbook(
        client=client, contractor=contractor, project=project,
        vessel=vessel, location=location, worksite=worksite,
        doc_ref=doc_ref, date=date, phases=phases,
        title_short="Snake Island"
    )
    return wb


# ==============================================================================
# PROJECT 2: SNEPCo Bonga IWS Mooring
# ==============================================================================
def gen_project_2():
    """SNEPCo Bonga IWS - Mooring Procedure for Jascon 65."""
    project = "Bonga IWS 2023 — Mooring Procedure"
    client = "SNEPCo (Shell Nigeria Exploration & Production Company)"
    contractor = "Atlantic Marine & Oilfield Services Ltd"
    vessel = "Jascon 65 (Light Diving Boat) / Jascon 66 (Mother Vessel)"
    location = "Bonga FPSO / SPM Buoy, Niger Delta, Nigeria"
    worksite = "Bonga FPSO and SPM Buoy — Port and Starboard Sides"
    doc_ref = "AM-SNEPCo-BONGA IWS-2023-PR-005 Rev.1"
    date = "25 07 2026"

    phases = []

    # Phase 1: Pre-Mobilisation
    phases.append({
        'name': 'Phase 1: Pre-Mobilisation & Planning',
        'tasks': [
            {
                'no': '1.1',
                'task': 'Pre-mobilisation of Jascon 65 and ROV equipment from Jascon 66',
                'hazard': '• Vessel-to-vessel transfer\n• Heavy ROV equipment lifting\n• Manual handling',
                'threat': 'Equipment drop during transfer between vessels.\nInadequate fender protection between vessels.\nRigger not qualified for vessel transfer operation.\nWind and current causing vessel separation.',
                'consequence': 'Equipment damage / personnel injury →\n• Dropped ROV equipment\n• Personnel overboard during transfer\n• Damage to Jascon 66 deck\n• Project delay',
                'controls': '1. Approved lift plan for all equipment transfers.\n2. Yokohama fenders positioned between vessels.\n3. IMCA LR006 — lifting guidelines compliance.\n4. Qualified riggers and banksman appointed.\n5. Manual handling — TILE assessment for heavy items.\n6. Pre-transfer meeting with both vessel captains.\n7. Dedicated comms channel between Jascon 65 and 66.',
                'recovery': 'Stop transfer operations.\nAssess damage to equipment.\nReplace if necessary.\nFirst aid for any injuries.\nReport to OIM/Client.',
                'action': 'Jascon 65 Capt. / Jascon 66 Capt. / PIC / Rigger / OCM',
                'p': '3C', 's': '2C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '1.2',
                'task': 'PTW preparation and risk assessment review',
                'hazard': '• Inadequate risk assessment\n• Permit gaps\n• SIMOPS with ongoing operations',
                'threat': 'Risk assessment not project-specific.\nPTW not raised or expired.\nConcurrent operations not identified.\nClient authority not consulted.',
                'consequence': 'Uncontrolled work →\n• Personnel exposed to unidentified hazards\n• Permit breach — regulatory non-compliance\n• SIMOPS incident\n• SNEPCo work stoppage',
                'controls': '1. Project-specific RA completed before PTW application.\n2. PTW raised and approved by Bonga Marine before operations.\n3. SIMOPS check — confirm no conflicting operations.\n4. Client authority review and sign-off on procedure.\n5. IMCA D014 — diving project plan requirements.\n6. Daily TBT conducted before each shift.\n7. PTW displayed in dive control and on vessel bridge.',
                'recovery': 'Suspend work until PTW valid.\nRaise revised PTW if scope changed.\nConsult Client Rep.\nReconduct TBT with all personnel.',
                'action': 'Diving Supervisor / HSE Lead / PIC / Client Rep / OIM',
                'p': '2C', 's': '3C', 'rp': '0', 'rs': '0',
            },
        ]
    })

    # Phase 2: Mooring Operations
    phases.append({
        'name': 'Phase 2: Mooring of Jascon 65 at Bonga FPSO',
        'tasks': [
            {
                'no': '2.1',
                'task': 'Yokohama fender deployment and connection',
                'hazard': '• Suspended load (fender)\n• Pinch points\n• Vessel movement',
                'threat': 'Fender dropped during deployment.\nFender connection fails under tension.\nJascon 65 moves unexpectedly during fender placement.',
                'consequence': 'Fender impact injury →\n• Crush injury to riggers\n• Equipment damage\n• Vessel hull damage\n• Delay to mooring operations',
                'controls': '1. Approved lift plan for fender deployment.\n2. Exclusion zone under suspended fender.\n3. Fender certified with valid SWL label.\n4. Riggers positioned clear of pinch points.\n5. Jascon 65 engines held off during fender deployment.\n6. Communication between Jascon 65 captain and FPSO riggers.\n7. IMCA LR006 — fender lifting guidelines.',
                'recovery': 'Stop fender deployment.\nAssess connection integrity.\nReposition fender if needed.\nReport to OIM if hull contact.',
                'action': 'Jascon 65 Capt. / Bonga Riggers / PIC / Diving Supervisor',
                'p': '3C', 's': '2C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '2.2',
                'task': 'Jascon 65 approach and mooring to FPSO',
                'hazard': '• Vessel collision\n• Mooring line snap-back\n• Compressed between vessel and FPSO\n• DP loss of position (if DP vessel)',
                'threat': 'Jascon 65 drifts into FPSO hull.\nMooring line parts — snap-back zone.\nCaptain misjudges vessel position.\nStrong current pushing vessel against FPSO.',
                'consequence': 'Vessel collision / mooring failure →\n• Hull damage to Jascon 65 and FPSO\n• Personnel caught between vessels\n• Mooring line injury (snap-back)\n• Diver in water — lost overboard\n• SNEPCo production impact',
                'controls': '1. Bonga Marine approval obtained before approach.\n2. Pre-approach assessment by Jascon 65 captain.\n3. Mooring lines prepared and coiled on FPSO deck.\n4. Constant radio comms on agreed channel.\n5. 2 riggers on FPSO deck to slack down mooring lines.\n6. Slack maintained on lines to avoid jarring.\n7. Spring line from weather side to control position.\n8. Weather window checked — wind <10 kts, current <1 knot.',
                'recovery': 'Stop approach immediately.\nReverse to safe distance.\nReassess weather and current.\nDeploy additional fenders if needed.\nRe-approach when conditions improve.',
                'action': 'Jascon 65 Capt. / ABs / Bonga Riggers / Marine Supervisor / OIM',
                'p': '3C', 's': '3C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '2.3',
                'task': 'LOTO implementation on Jascon 65 after mooring',
                'hazard': '• Residual energy sources\n• Electrical systems\n• Hydraulic systems',
                'threat': 'Engine started while LOTO active.\nElectrical panel not isolated.\nHydraulic equipment energised unexpectedly.\nLOTO tags removed by unauthorised person.',
                'consequence': 'Unexpected equipment activation →\n• Injury to divers or ROV crew\n• Damage to equipment\n• Loss of dive control\n• Vessel movement during dive operations',
                'controls': '1. LOTO procedure documented and approved.\n2. Captain, chief engineer, and ROV supervisor all sign LOTO.\n3. Keys handed to captain only after ROV de-isolation.\n4. LOTO tags displayed at engine room and electrical panel.\n5. Only authorised personnel may remove LOTO.\n6. LOTO log maintained — all entries recorded.\n7. LOTO checked every 6 hours during dive ops.',
                'recovery': 'Stop all diving and ROV operations.\nRe-implement LOTO.\nInvestigate breach.\nReport to OIM and Client Rep.\nReview procedure if recurring.',
                'action': 'Jascon 65 Capt. / Chief Engineer / ROV Supervisor / Diving Supervisor',
                'p': '2C', 's': '3C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '2.4',
                'task': 'Mooring at SPM Buoy',
                'hazard': '• SPM Buoy turntable rotation\n• Mooring line tension\n• Confined space on buoy deck',
                'threat': 'Turntable not locked — buoy rotates during mooring.\nMooring line over-tensioned.\nRigger falls from buoy deck.\nCurrent exceeds safe limit for buoy mooring.',
                'consequence': 'Buoy damage / personnel injury →\n• Rigger injury from fall\n• Mooring line failure\n• SPM Buoy structural damage\n• Dislodgement of subsea equipment',
                'controls': '1. SPM turntable locked by Bonga Marine crew before approach.\n2. Bonga Marine clearance obtained from Captain.\n3. 2 riggers stationed on SPM Buoy.\n4. Mooring lines tensioned appropriately — not over-tensioned.\n5. Fall protection worn by riggers on buoy.\n6. Current monitoring — mooring suspended if >1 knot.\n7. IMCA D067 — current effects on diver and vessel safety.',
                'recovery': 'Stop mooring.\nRelease lines safely.\nRe-lock turntable.\nReassess conditions.\nRe-attempt when safe.',
                'action': 'Jascon 65 Capt. / SPM Buoy Riggers / Marine Supervisor / Bonga Marine',
                'p': '3C', 's': '3C', 'rp': '0', 'rs': '0',
            },
        ]
    })

    # Phase 3: ROV Operations with Moored Vessel
    phases.append({
        'name': 'Phase 3: ROV Operations — In-Water Survey',
        'tasks': [
            {
                'no': '3.1',
                'task': 'ROV deployment and survey operations at Bonga FPSO',
                'hazard': '• ROV deployment from LARS/basket\n• Electrical equipment in wet environment\n• Umbilical entanglement\n• Underwater obstacles',
                'threat': 'ROV dropped during deployment.\nUmbilical snagged on FPSO structure.\nROV camera obstruction by marine growth.\nPower failure during survey.\nVessel movement causing ROV cable stress.',
                'consequence': 'ROV damage / survey data loss →\n• ROV loss or damage\n• Umbilical severed\n• Survey of flexjoints incomplete\n• Schedule impact\n• Cost of ROV repair/replacement',
                'controls': '1. ROV PMS and certification current.\n2. ROV deployment plan approved by ROV Supervisor.\n3. Umbilical managed — tender monitors tension.\n4. 3m karabiner maintaining ROV proximity to swim line.\n5. Pre-deployment system check and test.\n6. Dedicated power source for ROV — isolated from other ops.\n7. Communication test before deployment.\n8. Marine growth assessment before dive operations.',
                'recovery': 'Terminate ROV operations.\nRecover ROV to deck.\nRepair damaged umbilical.\nDMT assessment.\nRe-deploy when fixed.',
                'action': 'ROV Supervisor / ROV Pilot / Technician / Diving Supervisor / LARS Operator',
                'p': '3C', 's': '2C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '3.2',
                'task': 'ROV inspection of FPSO hull and SPM Buoy structure',
                'hazard': '• Confined space access\n• Marine growth / sharp edges\n• Underwater current\n• Diver/ROV SIMOPS',
                'threat': 'Diver and ROV operating in same zone.\nDiver reaches into confined space between hull and structure.\nCurrent carrying diver into ROV umbilical.\nMarine growth obscuring inspection targets.',
                'consequence': 'SIMOPS incident →\n• ROV umbilical entangled with diver\n• Diver struck by ROV\n• Structural damage to hull from ROV contact\n• Survey data compromised',
                'controls': '1. IMCA D054 — combined ROV/diving coordination.\n2. Separate dive and ROV zones where possible.\n3. Dual responsibility — Diving Supt and ROV Supervisor coordinate.\n4. Diver D-ring demarcation from ROV operating area.\n5. Underwater comms test before operations.\n6. Current monitoring per IMCA D067.\n7. Marine growth removal if obscuring inspection targets.',
                'recovery': 'Stop all operations.\nSeparate diver and ROV.\nAssess damage.\nRe-plan zones.\nResume when safe.',
                'action': 'ROV Supervisor / Diving Supervisor / ROV Pilot / Tender / Diver',
                'p': '3C', 's': '3C', 'rp': '0', 'rs': '0',
            },
        ]
    })

    # Phase 4: Demobilisation
    phases.append({
        'name': 'Phase 4: Unmooring & Demobilisation',
        'tasks': [
            {
                'no': '4.1',
                'task': 'Unmooring of Jascon 65 from FPSO / SPM Buoy',
                'hazard': '• Mooring line tension release\n• Vessel movement\n• Underwater line entanglement',
                'threat': 'Mooring line parting — snap-back.\nLine wrapped around propeller.\nVessel drifts during unmooring.\nLOTO not properly removed.',
                'consequence': 'Injury during unmooring →\n• Rigger struck by snapping line\n• Vessel damage\n• Propeller damage from entangled line\n• Vessel collision',
                'controls': '1. ROV recovered and system powered down before unmooring.\n2. LOTO removal by Captain, Chief Engineer, ROV Supervisor.\n3. Bonga Marine unlock SPM turntable (if at buoy).\n4. Mooring lines paid out gradually.\n5. Captain controls vessel movement — speed kept low.\n6. ABs at bow and stern to receive lines.\n7. Good housekeeping on deck.',
                'recovery': 'Stop unmooring.\nSecure lines.\nReassess plan.\nReport to OIM if damage sustained.',
                'action': 'Jascon 65 Capt. / ABs / Bonga Riggers / Chief Engineer / ROV Supervisor',
                'p': '3C', 's': '2C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '4.2',
                'task': 'Transit back to base / next dive location',
                'hazard': '• Vessel transit in sea state\n• Weather deterioration\n• Navigation hazard',
                'threat': 'Swell exceeds safe transit limit.\nVisibility reduced by weather.\nVessel navigation error in field.',
                'consequence': 'Transit incident →\n• Vessel damage from wave action\n• Personnel injury on deck\n• Navigation hazard\n• Equipment loss overboard',
                'controls': '1. Sea state check — swell <1.5m for DSV transit.\n2. Weather forecast reviewed before departure.\n3. Navigation check — GPS and chart plotter operational.\n4. All loose equipment secured.\n5. Vessel Master approval for transit.\n6. Engine check before departure.\n7. Daily progress report filed.',
                'recovery': 'Divert to safe harbour.\nReport to OIM.\nAwait weather improvement.\nAssess equipment damage.',
                'action': 'Vessel Master / Jascon 65 Capt. / Marine Superintendent / OIM',
                'p': '2C', 's': '2C', 'rp': '0', 'rs': '0',
            },
        ]
    })

    wb = create_hira_workbook(
        client=client, contractor=contractor, project=project,
        vessel=vessel, location=location, worksite=worksite,
        doc_ref=doc_ref, date=date, phases=phases,
        title_short="Bonga IWS"
    )
    return wb


# ==============================================================================
# PROJECT 3: 3D Photogrammetry Inspection
# ==============================================================================
def gen_project_3():
    """SNEPCo Bonga Pipelines - 3D Photogrammetry Inspection of Flexjoints."""
    project = "Bonga Pipelines Flexjoints — 3D Photogrammetry Inspection"
    client = "SNEPCo (Shell Nigeria Exploration & Production Company)"
    contractor = "West African Ventures (WAV)"
    vessel = "Diving Support Vessel (DSV) / FSU"
    location = "Bonga Field, Niger Delta — 1200m water depth, 120km SW of Niger Delta"
    worksite = "Bonga FPSO — 11 Flexjoints (10\" PFL-01/02/08/09/11/12, 10\" WFL-01, 12\" PFL-05/06, 12\" WFL-03, 16\" Gas Export)"
    doc_ref = "BON-FLEX-SNEPCO-VA-5793-00016-000 Rev.A01"
    date = "25 07 2026"

    phases = []

    # Phase 1: Pre-Mobilisation
    phases.append({
        'name': 'Phase 1: Pre-Mobilisation & Planning',
        'tasks': [
            {
                'no': '1.1',
                'task': 'Mobilisation of photogrammetry equipment and diving team',
                'hazard': '• Heavy photogrammetry camera systems\n• Lifting operations\n• Equipment transfer between vessels',
                'threat': 'Camera system dropped during loading.\nLifting gear certificate expired.\nEquipment not secured for vessel transit.',
                'consequence': 'Equipment damage / injury →\n• 3D camera system damage\n• Personnel injury during handling\n• Project delay\n• Cost of camera replacement',
                'controls': '1. Approved lift plan for camera system loading.\n2. Certified lifting gear with valid colour code.\n3. IMCA LR006 — lifting guidelines.\n4. Equipment crated and secured for transit.\n5. Pre-transfer inspection of all photogrammetry equipment.\n6. Manual handling — TILE assessment for heavy items.\n7. Dedicated equipment handler appointed.',
                'recovery': 'Stop lifting operations.\nAssess damage to equipment.\nReplace with backup camera.\nReport to Client Project Manager.',
                'action': 'PIC / Rigger / Crane Operator / Diving Supervisor / OCM',
                'p': '3C', 's': '2C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '1.2',
                'task': 'Inspection planning — flexjoint selection and prioritisation',
                'hazard': '• Inadequate inspection scope\n• Missed flexjoints\n• Hydrocarbon exposure',
                'threat': 'Not all 11 flexjoints scheduled for inspection.\nFlexjoint identified as high-risk but not prioritised.\nHydrocarbon residue on flexjoint elastomer — vapour release.',
                'consequence': 'Incomplete inspection →\n• Missed anomalies on high-risk flexjoints\n• Elasto mer degradation undetected\n• Production flowline failure\n• Environmental spill',
                'controls': '1. All 11 flexjoints listed and cross-checked against scope.\n2. High-risk flexjoints prioritised per SNEPCo criteria.\n3. Hydrocarbon hazard assessment for each flexjoint.\n4. Inspection sequence planned to minimise vessel repositioning.\n5. 3D mapping equipment tested before deployment.\n6. Photo/video documentation protocol established.\n7. Client review and approval of inspection plan.',
                'recovery': 'Update inspection plan.\nDeploy additional diver/ROV.\nRe-prioritise flexjoints.\nReport to Client Project Manager.',
                'action': 'Project Engineer / Diving Supervisor / Client Rep / ROV Supervisor',
                'p': '2C', 's': '3C', 'rp': '0', 'rs': '0',
            },
        ]
    })

    # Phase 2: Subsea Inspection
    phases.append({
        'name': 'Phase 2: Subsea Inspection of Flexjoints',
        'tasks': [
            {
                'no': '2.1',
                'task': 'Diver-assisted visual inspection and photography of flexjoints',
                'hazard': '• Hydrocarbon exposure on flexjoint surface\n• Sharp edges on flexjoint casing\n• Underwater current at 1200m (deep water)\n• Confined space between riser and FPSO hull\n• Marine growth on flexjoint\n• Fatigue from extended bottom time',
                'threat': 'Diver touching hydrocarbon-contaminated elastomer.\nDiver reaching into tight space between riser and hull.\nCurrent exceeding 1 knot at depth.\nDiver positioned between flexjoint and vessel structure.\nMarine growth obscuring anomaly detection.\nExceeding 4 hours in-water per 24 hours.',
                'consequence': 'Diver injury / inspection failure →\n• Chemical exposure from hydrocarbon residue\n• Crush injury in confined space\n• DCI from deep water dive\n• Umbilical entanglement\n• Missed anomalies — incomplete photographic record\n• Production flowline failure undetected',
                'controls': '1. IMCA D010 — diving from DP vessels protocol.\n2. IMCA D078 — umbilical management (5m restraint, D-ring demarcation).\n3. PPE — chemical-resistant gloves and full coverage suit.\n4. Underwater comms test before each dive.\n5. Current monitoring — dive suspended if >1 knot (IMCA D067).\n6. 3D photogrammetry camera system tested pre-dive.\n7. Diver tracked via USBL position tracking.\n8. Pre-dive briefing on all flexjoint locations and hazards.\n9. Maximum 4 hours in-water per 24 hours.\n10. Standby diver in immediate readiness.\n11. IMCA D058 — weak link for diver attachment.',
                'recovery': 'Emergency basket recovery via LARS.\nStandby diver deployed.\nChemical decontamination if hydrocarbon exposure.\nDMT assessment.\nDDC on 18-hour bend watch.\nERP — medevac if serious.\nRe-inspect flexjoint when safe.',
                'action': 'Diving Supervisor / Tender / Diver / Standby Diver / Diver Medic / OCM',
                'p': '3D', 's': '3C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '2.2',
                'task': '3D mapping and anomaly documentation',
                'hazard': '• Underwater camera handling\n• Electrical equipment in wet environment\n• Diver distraction from camera work',
                'threat': 'Camera cable tangled in umbilical.\nElectrical fault on camera system.\nDiver focused on camera positioning — missing hazard.\nPoor light conditions reducing photo quality.',
                'consequence': 'Incomplete data / equipment damage →\n• 3D model incomplete — anomalies missed\n• Camera system water damage\n• Electrical fault — shock risk to diver\n• Repeated dives required — increased DCI risk',
                'controls': '1. Camera system waterproof rating verified.\n2. Electrical isolation for all underwater equipment.\n3. Dedicated diver for camera operation — not tasking diver with other work.\n4. Underwater lighting setup for clear photography.\n5. Backup camera system on standby.\n6. Photo quality check surface-side before diver recovery.\n7. IMCA D018 — examination of diving plant.',
                'recovery': 'Retrieve camera system.\nDry and inspect equipment.\nReplace if damaged.\nRe-deploy backup camera.\nAssess if data quality acceptable.',
                'action': 'Photogrammetry Technician / Diving Supervisor / Diver / Diver Medic',
                'p': '2C', 's': '2C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '2.3',
                'task': 'ROV-assisted inspection of flexjoints (alternative/backup)',
                'hazard': '• ROV deployment from vessel\n• Umbilical entanglement\n• ROV contact with flexjoint',
                'threat': 'ROV dropped during deployment.\nUmbilical snagged on FPSO structure.\nROV thruster damaging flexjoint elastomer.\nPower failure during inspection.',
                'consequence': 'ROV damage / flexjoint damage →\n• ROV loss or damage\n• Umbilical severed\n• Flexjoint elastomer damaged by ROV contact\n• Survey data loss\n• Schedule impact',
                'controls': '1. ROV PMS and certification current.\n2. ROV deployment plan approved by ROV Supervisor.\n3. ROV thruster guard installed.\n4. Umbilical tendering — monitor tension.\n5. ROV speed limited near flexjoints.\n6. Pre-deployment system check.\n7. IMCA D054 — ROV/diving coordination.\n8. Dedicated comms channel.',
                'recovery': 'Terminate ROV operations.\nRecover ROV.\nRepair damage.\nRe-deploy when fixed.',
                'action': 'ROV Supervisor / ROV Pilot / Technician / Diving Supervisor / LARS Operator',
                'p': '3C', 's': '2C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '2.4',
                'task': 'Post-inspection assessment — anomaly evaluation and reporting',
                'hazard': '• Data processing error\n• Anomaly misclassification\n• Report delay',
                'threat': '3D model processing error — false anomaly.\nDegraded elastomer missed in photo analysis.\nReport not submitted to SNEPCo on time.\nData corrupted during transfer.',
                'consequence': 'Incorrect assessment →\n• False alarm — unnecessary production shutdown\n• Missed critical degradation\n• Flexjoint failure in service\n• Production loss\n• Regulatory non-compliance',
                'controls': '1. 3D model reviewed by qualified pipeline engineer.\n2. Anomaly classified per SNEPCo acceptance criteria.\n3. Cross-check with previous inspection data.\n4. Report template approved by SNEPCo.\n5. Data backup — redundant storage.\n6. Report submitted within agreed timeline.\n7. QA/QC review before client delivery.',
                'recovery': 'Re-process 3D data.\nRe-inspect if anomalies uncertain.\nEscalate to SNEPCo Pipeline Engineer.\nSchedule follow-up inspection if needed.',
                'action': 'Pipeline Engineer / Project Engineer / Client Rep / HSE Lead',
                'p': '2C', 's': '3C', 'rp': '0', 'rs': '0',
            },
        ]
    })

    # Phase 3: Close-out
    phases.append({
        'name': 'Phase 3: Demobilisation & Reporting',
        'tasks': [
            {
                'no': '3.1',
                'task': 'Demobilisation of photogrammetry equipment',
                'hazard': '• Lifting operations\n• Equipment handling\n• Vessel departure',
                'threat': 'Camera system dropped during unloading.\nEquipment not secured for transit.\nVessel departure in poor sea state.',
                'consequence': 'Equipment damage →\n• 3D camera system damage\n• Project delay\n• Cost of camera repair\n• Personnel injury',
                'controls': '1. Approved demob lift plan.\n2. Equipment crated before unloading.\n3. Sea state check — swell <1.5m.\n4. All equipment inventory checked.\n5. Vessel Master approval for departure.\n6. IMCA LR006 — lifting during transit.\n7. Post-demob equipment condition report.',
                'recovery': 'Stop demob.\nSecure equipment.\nReassess conditions.\nReport to Client.',
                'action': 'Vessel Master / PIC / Rigger / Crane Operator / Diving Supervisor',
                'p': '2C', 's': '2C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '3.2',
                'task': 'Final inspection report submission to SNEPCo',
                'hazard': '• Data integrity\n• Confidentiality breach\n• Report accuracy',
                'threat': 'Confidential SNEPCo data leaked.\nInspection report contains errors.\nFlexjoint anomalies not properly classified.\nReport not delivered to agreed timeline.',
                'consequence': 'Data breach / incorrect assessment →\n• SNEPCO confidential information compromised\n• Incorrect flexjoint condition assessment\n• Unnecessary production impact\n• Regulatory non-compliance\n• Contractual penalties',
                'controls': '1. Report reviewed by qualified engineer before submission.\n2. Confidentiality agreement with SNEPCo maintained.\n3. All 11 flexjoints accounted for in report.\n4. Anomaly classification per SNEPCo criteria.\n5. 3D maps and photos attached to report.\n6. Report submitted to designated SNEPCo contact.\n7. Copy filed with project documentation.',
                'recovery': 'Correct report errors.\nRe-submit to SNEPCo.\nNotify Client Rep.\nAssess contractual impact.',
                'action': 'Project Engineer / Pipeline Engineer / Client Rep / HSE Lead',
                'p': '2C', 's': '2C', 'rp': '0', 'rs': '0',
            },
        ]
    })

    wb = create_hira_workbook(
        client=client, contractor=contractor, project=project,
        vessel=vessel, location=location, worksite=worksite,
        doc_ref=doc_ref, date=date, phases=phases,
        title_short="Bonga Flexjoints"
    )
    return wb


# ==============================================================================
# PROJECT 4: Flange Disconnection/Reconnection (SAL Buoy)
# ==============================================================================
def gen_project_4():
    """Atlantic Marine - SAL Buoy Riser Flange Disconnection and Reconnection."""
    project = "Single Anchor Loading (SAL) Buoy — Riser Flange Disconnection & Reconnection"
    client = "Atlantic Marine & Oilfield Services Ltd"
    contractor = "ConsultingSubsea / Approved Diving Contractor"
    vessel = "African Concept (DP Vessel)"
    location = "Bonga Field / Offshore Nigeria"
    worksite = "SAL (Single Anchor Loading) Buoy Riser"
    doc_ref = "AM-SAL-AFR-HIRA-2026"
    date = "25 07 2026"

    phases = []

    # Phase 1: Pre-Dive Preparation
    phases.append({
        'name': 'Phase 1: Pre-Mobilisation & Preparation',
        'tasks': [
            {
                'no': '1.1',
                'task': 'SAL Buoy isolation confirmation and PTW verification',
                'hazard': '• Incomplete isolation\n• Residual pressure in riser\n• SIMOPS with production',
                'threat': 'Isolation certificates not signed off.\nRiser still pressurised — hydrocarbon release.\nConcurrent operations not identified.\nClient authority not consulted before diving.',
                'consequence': 'Uncontrolled hydrocarbon release →\n• Fire / explosion hazard\n• Diver fatality\n• FPSO production loss\n• Environmental damage\n• Full ERP activation — critical incident',
                'controls': '1. SAL Buoy isolations confirmed by designated Client representative.\n2. LOTO certificates verified and signed off — Diving Supervisor in attendance.\n3. PTW valid with no SIMOPS taking place.\n4. Client authority approval of procedure before diving commences.\n5. IMCA D010 — DP diving protocol.\n6. Daily TBT conducted — all threats and controls reviewed.\n7. Alpha flag raised — marine channel broadcast for diving ops.',
                'recovery': 'Stop all diving operations.\nVerify isolations.\nRaise revised PTW.\nConsult Client Rep.\nReconduct TBT.',
                'action': 'Diving Supervisor / Client Rep / PIC / Marine Superintendent / OCM',
                'p': '2D', 's': '5C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '1.2',
                'task': 'Diving equipment inspection and pre-dive checks',
                'hazard': '• Defective dive equipment\n• Contaminated air supply\n• Electrical faults',
                'threat': 'Compressor oil carryover.\nDamaged umbilical from previous dive.\nFaulty electrical insulation.\nBreathing air analysis not current.',
                'consequence': 'Equipment failure →\n• Loss of breathing gas to diver\n• Electric shock underwater\n• Contaminated air — CO poisoning\n• DCI requiring recompression\n• Loss of diver communication',
                'controls': '1. IMCA D018 — examination and testing of diving plant — PMS current.\n2. Air analysis per BS EN 12012 — O2, CO, CO2 within limits.\n3. All bail-out cylinders checked with calibrated Go/No-Go gauges.\n4. IMCA D043 — gas cylinder marking and certification.\n5. IMCA D022/D023/D040 DESIGN audit current.\n6. Pre-dive checklist completed and signed — control van, deck, chamber, divers.\n7. DDC prepped and ready — 18-hour bend watch capability.',
                'recovery': 'Terminate dive.\nReplace defective equipment.\nRe-test air supply.\nDMT assessment before resuming.\nNotify Diving Superintendent.',
                'action': 'Dive Technician / Diving Supervisor / Diver Medic / Tender',
                'p': '2C', 's': '3C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '1.3',
                'task': 'DP vessel setup and positioning',
                'hazard': '• DP loss of position\n• Thruster proximity to diver\n• PRS dropout',
                'threat': 'DP software failure or PRS dropout.\nThruster drawing diver towards vessel.\nVessel drifting due to current exceeding DP capacity.\nLoss of position — drive-off or drift-off scenario.',
                'consequence': 'DP failure →\n• Diver struck by vessel hull or thruster\n• Umbilical severed by vessel movement\n• Vessel collision with SAL buoy\n• Riser dislodged from structure\n• Full ERP activation — critical incident',
                'controls': '1. DP reliability programme — trained DP operators, ASOG, FMEA.\n2. IMCA D010 — DP status comms protocol to Dive Supervisor.\n3. IMCA M190 — annual DP trials current.\n4. IMCA M273 — DP drills conducted before diving ops.\n5. IMCA M117 — DP personnel training and competency.\n6. DP alarms at dive control, bridge, and dive supervision.\n7. D-ring demarcation system per vessel excursion plan.\n8. Active umbilical tending — tender monitors payout.',
                'recovery': 'DP alarm activated — all divers surface immediately.\nEmergency basket recovery.\nAssess DP status.\nDMT assessment.\nERP Tier 2 activation.\nResume only when DP confirmed stable.',
                'action': 'Diving Supervisor / DP Operator / Vessel Master / Tender / Standby Diver',
                'p': '3C', 's': '4C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '1.4',
                'task': 'Pre-dive briefing and weather window assessment',
                'hazard': '• Inadequate briefing\n• Unsuitable weather\n• Fatigued personnel',
                'threat': 'Divers not fully briefed on hazards and controls.\nWeather deteriorating — wind >10 kts or swell >1.5m.\nDiver exceeding working hours — fatigue.\nComms protocol not confirmed.',
                'consequence': 'Uncontrolled dive →\n• Diver unaware of hazards\n• Operations suspended mid-dive\n• DCI from fatigue-induced error\n• Communication failure underwater\n• Emergency response compromised',
                'controls': '1. Pre-dive briefing completed — all divers and standby diver present.\n2. Weather window checked — wind <10 kts, current <1 knot, swell <1.5m.\n3. IMCA D067 — underwater currents and diver safety guidelines.\n4. Working hours monitored — max 12 hours continuous.\n5. Comms test — diver to supervisor, bridge, and dive control.\n6. All personnel aware of communication protocols.\n7. Diver competency confirmed for SAL flange task.',
                'recovery': 'Delay dive until weather improves.\nBrief all personnel.\nRelieve fatigued diver.\nRe-test comms.\nResume when all clear.',
                'action': 'Diving Supervisor / All Diving Team / Vessel Master / Bridge',
                'p': '2C', 's': '2C', 'rp': '0', 'rs': '0',
            },
        ]
    })

    # Phase 2: Diver Deployment & Excursion
    phases.append({
        'name': 'Phase 2: Diver Deployment & Excursion to SAL Buoy',
        'tasks': [
            {
                'no': '2.1',
                'task': 'Diver deployment via LARS basket to SAL Buoy work site',
                'hazard': '• LARS winch wire / A-frame failure\n• Basket snag on vessel hull\n• Uncontrolled descent rate\n• Umbilical entanglement\n• Diver DCI from rapid descent',
                'threat': 'LARS wire parted — shock load or fatigue.\nBasket snagging on hull during descent.\nWinch brake failure — too rapid descent.\nUmbilical trailing line caught on hull.\nBasket descent too fast — exceeding 30 fsw/min ascent rate.',
                'consequence': 'Deployment failure →\n• Diver struck by basket or hull\n• Umbilical severed\n• Diver DCI from uncontrolled descent\n• LARS equipment damage\n• Full ERP activation',
                'controls': '1. LARS PMS and D023 certification current.\n2. IMCA D010 — DP diving deployment protocol.\n3. IMCA D078 — umbilical management (5m restraint, D-ring demarcation).\n4. Pre-lower test — basket lowered empty to working depth.\n5. Diver tracked via Mesotech survey beacon.\n6. Winch operator certified and experienced.\n7. Basket descent rate monitored — max 30 fsw/min.\n8. Active umbilical tending from surface.\n9. Standby diver dressed and in immediate readiness.',
                'recovery': 'Stop descent immediately.\nSecure basket.\nAssess diver status.\nEmergency recovery if needed.\nDMT assessment.',
                'action': 'LARS Operator / Diving Supervisor / Tender / Diver / Standby Diver',
                'p': '3C', 's': '3C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '2.2',
                'task': 'Diver transit to SAL Buoy — umbilical management',
                'hazard': '• Umbilical entanglement on transit\n• Vessel thruster proximity\n• Strong underwater current\n• Poor visibility',
                'threat': 'Umbilical caught on vessel structure during transit.\nDiver carried by current toward thruster zone.\nSilt-out reducing visibility to zero.\nDiver exceeding excursion limit beyond D-ring demarcation.',
                'consequence': 'Umbilical damage →\n• Severed umbilical — loss of gas supply\n• Diver stranded underwater\n• DCI from emergency response\n• Vessel thruster damage to diver\n• ERP Tier 2 activation',
                'controls': '1. IMCA D078 — umbilical marked at 5m/10m intervals (red/black tape).\n2. Working diver umbilical restrained — cannot reach within 5m of thruster.\n3. Standby diver umbilical restrained at 3m from hazard, 2m beyond working diver.\n4. 3m karabiner maintaining diver proximity to swim line — lockable type.\n5. Vessel hazard diagram on bridge and in dive control per D078 §10.\n6. Current monitoring — dive suspended if >1 knot.\n7. Underwater lights for low visibility transit.\n8. USBL position tracking of diver.',
                'recovery': 'Stop diver transit.\nRe-route umbilical.\nDeploy standby diver if entangled.\nSurface if umbilical severed.\nEmergency decompression if required.',
                'action': 'Diving Supervisor / Tender / Diver / Standby Diver / LARS Operator',
                'p': '3C', 's': '3C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '2.3',
                'task': 'Diver rigging installation on SAL Riser',
                'hazard': '• Sharp edges on riser\n• Confined space between riser and buoy\n• Manual handling underwater\n• Diver fatigue / extended bottom time',
                'threat': 'Diver working in confined space.\nHMPE slings damaged by sharp edges.\nDiver reaching beyond excursion limit.\nBolts corroded and seized — excessive force required.',
                'consequence': 'Diver injury during rigging →\n• Crush injury from shifting riser\n• Laceration from sharp edges\n• Umbilical snagged on riser\n• DCI from extended bottom time\n• Dropped load / uncontrolled swing',
                'controls': '1. IMCA D078 — D-ring demarcation system confirmed.\n2. Edge protection (sleeving) between riser and HMPE slings.\n3. 2 x 1m 1ton HMPE slings — certified and colour coded.\n4. Active umbilical tending throughout.\n5. Diver work pace managed — rest breaks as needed.\n6. Continuous comms with supervisor.\n7. IMCA D067 — diver fatigue and current effects.\n8. Pre-rigging diver briefing.',
                'recovery': 'Stop rigging.\nSurface diver if injury suspected.\nDMT assessment.\nReassess rigging approach.\nResume when diver fit.',
                'action': 'Diving Supervisor / Tender / Diver / Standby Diver',
                'p': '3C', 's': '3C', 'rp': '0', 'rs': '0',
            },
        ]
    })

    # Phase 3: Flange Work
    phases.append({
        'name': 'Phase 3: Flange Disconnection — Riser Recovery',
        'tasks': [
            {
                'no': '3.1',
                'task': 'Diver loosens and removes flange bolts (8 bolts on riser flange)',
                'hazard': '• Hand tool injury (pinch points, impact)\n• Defective tools (cracked flogging spanner)\n• Dropped objects (bolts, spanner, hammer)\n• Corroded/seized bolts requiring excessive force',
                'threat': 'Spanner slips — hand injury to diver.\nHammer head loose — detached during use.\nBolts dropped on diver or equipment.\nBolts seized — excessive torque causing tool failure.\nDiver fatigued from extended bottom time.',
                'consequence': 'Hand injury / dropped tool →\n• Diver finger/hand crush injury\n• Tool dropped — damaged subsea equipment\n• Bolt recovered to bag\n• DCI from emergency response\n• Flange disconnection delayed',
                'controls': '1. All tools pre-dive inspected — spanners, hammers, bolts bags.\n2. Defective tools marked DO NOT USE.\n3. Tool lanyards — all tools tethered to diver.\n4. IMCA D018 — examination of diving plant and tools.\n5. Bolts recovered to designated bolts bag.\n6. Diver work pace managed.\n7. Continuous comms with supervisor.\n8. Hand protection — cut-resistant gloves.\n9. IMCA D078 — active umbilical tending.',
                'recovery': 'Stop work.\nAssess diver hand.\nFirst aid by vessel medic.\nRecover diver if serious injury.\nReplace defective tools.\nResume when safe.',
                'action': 'Diving Supervisor / Tender / Diver / Standby Diver / Diver Medic',
                'p': '3C', 's': '2C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '3.2',
                'task': 'Diver attaches crane stinger to rigging on riser',
                'hazard': '• Suspended load\n• Underwater rigging\n• Crane stinger handling',
                'threat': 'Crane stinger not correctly attached.\nRigging not balanced — riser swings.\nCrane operator loses sight of diver.\nRiser weight exceeds estimated.',
                'consequence': 'Riser impact →\n• Diver struck by swinging riser\n• Riser drops — damage to buoy structure\n• Umbilical severed\n• Personnel injury on deck\n• Buoy structural damage',
                'controls': '1. Approved lift plan for riser recovery.\n2. Crane stinger tested and certified.\n3. 2 x 1m 1ton HMPE slings — balanced.\n4. Dedicated crane signalman with comms to operator.\n5. IMCA LR006 — lifting guidelines.\n6. Pre-lift meeting with all parties.\n7. Trial lift to 0.5m to verify stability.\n8. SWL clearly marked on all gear.',
                'recovery': 'Stop crane operation.\nRecover diver to basket.\nReassess rigging.\nRevised lift plan if needed.\nResume when stable.',
                'action': 'Diving Supervisor / Crane Operator / Diver / Tender / Deck Crew',
                'p': '3C', 's': '3C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '3.3',
                'task': 'Removal of remaining bolts and riser recovery to surface',
                'hazard': '• Suspended load\n• Dropped bolts\n• Confined space extraction',
                'threat': 'Bolts dropped overboard.\nRiser contacts buoy structure during recovery.\nDiver trapped between riser and buoy.\nCrane recovery speed too fast.',
                'consequence': 'Riser recovery incident →\n• Diver crushed between riser and buoy\n• Buoy structural damage\n• Dropped bolts — environmental contamination\n• Umbilical damage\n• Project delay',
                'controls': '1. Bolts recovered to bag — no dropping overboard.\n2. Crane recovery at controlled speed.\n3. Diver maintains distance from riser during lift.\n4. Crane operator follows signalman instructions.\n5. Exclusion zone under suspended load.\n6. IMCA LR006 — lift plan compliance.\n7. Trial lift verified before full recovery.\n8. Diver clears riser area before crane lift begins.',
                'recovery': 'Stop crane operation.\nSecure riser on deck.\nCheck diver for injury.\nInspect buoy for damage.\nResume when safe.',
                'action': 'Crane Operator / Diving Supervisor / Diver / Tender / Deck Crew / OCM',
                'p': '3C', 's': '3C', 'rp': '0', 'rs': '0',
            },
        ]
    })

    # Phase 4: Riser Deployment & Reconnection
    phases.append({
        'name': 'Phase 4: Riser Deployment & Flange Reconnection',
        'tasks': [
            {
                'no': '4.1',
                'task': 'Crane recovery of riser to vessel deck',
                'hazard': '• Suspended load\n• Congested deck\n• Manual handling on deck',
                'threat': 'Riser swings uncontrollably.\nRiser contacts vessel superstructure.\nRiggers not in exclusion zone.\nDeck crew not positioned safely.',
                'consequence': 'Riser impact with vessel →\n• Vessel hull damage\n• Dropped load\n• Personnel injury on deck\n• Equipment damage\n• Production delay',
                'controls': '1. Exclusion zone under suspended load.\n2. Tag line deployed to control swing.\n3. Crane operator and signalman on dedicated comms.\n4. Riggers positioned safely on deck.\n5. Deck clear of obstructions.\n6. IMCA LR006 — approved lift plan.\n7. Riser landed on deck with chocks and supports.\n8. Good housekeeping on deck.',
                'recovery': 'Stop crane operation.\nReposition deck crew.\nReassess lift plan.\nResume when clear.',
                'action': 'Crane Operator / Signalman / Deck Crew / PIC / Diving Supervisor',
                'p': '3C', 's': '2C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '4.2',
                'task': 'Riser re-deployment — diver assisted alignment and flange connection',
                'hazard': '• Suspended load\n• Sharp edges on flange\n• Confined space\n• Hand tool injury\n• Poor visibility',
                'threat': 'Riser misaligned — flange faces not mating.\nDiver reaches between flange faces.\nBolts seized — excessive torque.\nRiser shifts during bolting.\nSilt-out reducing visibility.',
                'consequence': 'Flange connection failure →\n• Diver crushed between flanges\n• Riser dropped — damage to buoy\n• Hydrocarbon release if isolation inadequate\n• Project delay\n• Bolting not completed to specification',
                'controls': '1. Approved lift plan for riser re-deployment.\n2. Riser aligned by crane before diver engagement.\n3. Diver maintains safe distance from flange faces.\n4. Bolting sequence — cross-tightening pattern.\n5. Torque wrench used for final tightening.\n6. IMCA D078 — umbilical management.\n7. Continuous comms with supervisor.\n8. Edge protection on riser.\n9. Underwater lights for visibility.\n10. Pre-bolting diver briefing on flange hazards.',
                'recovery': 'Stop work.\nSurface diver if injury.\nRealign riser if needed.\nDMT assessment.\nResume bolting when safe.',
                'action': 'Diving Supervisor / Diver / Crane Operator / Tender / Standby Diver',
                'p': '3C', 's': '3C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '4.3',
                'task': 'Post-connection verification and isolation removal',
                'hazard': '• Residual pressure\n• Hydrocarbon release\n• Pressure testing',
                'threat': 'Isolation not fully removed before pressurisation.\nFlange leak — hydrocarbon release.\nPressure test exceeds design limits.\nBolts not torqued to specification.',
                'consequence': 'Hydrocarbon release →\n• Fire / explosion\n• Diver fatality\n• Environmental damage\n• FPSO production impact\n• Regulatory penalties',
                'controls': '1. LOTO certificates verified before isolation removal.\n2. Flange leak test — pressure test per design specification.\n3. All bolts torqued and verified.\n4. Pressure test monitored by Diving Supervisor.\n5. Emergency response equipment ready.\n6. IMCA D010 — DP diving protocol.\n7. Client authority present for pressurisation.\n8. Post-connection inspection by diver.\n9. Photo documentation of completed connection.',
                'recovery': 'Stop pressurisation.\nDepressurise safely.\nRepair flange leak.\nRe-torque bolts.\nRe-test.\nNotify Client Rep.',
                'action': 'Diving Supervisor / Client Rep / PIC / Vessel Master / Diver / Deck Crew',
                'p': '2D', 's': '5C', 'rp': '0', 'rs': '0',
            },
        ]
    })

    # Phase 5: Post-Dive Close-out
    phases.append({
        'name': 'Phase 5: Post-Dive Close-out & Recovery',
        'tasks': [
            {
                'no': '5.1',
                'task': 'Diver recovery to basket and ascent',
                'hazard': '• Uncontrolled ascent rate\n• DCI from rapid ascent\n• Umbilical entanglement during recovery',
                'threat': 'Diver ascending too fast — exceeding 30 fsw/min.\nUmbilical snagged on structure during ascent.\nDiver exceeds safe bottom time per dive tables.\nDecompression stops missed.',
                'consequence': 'DCI from uncontrolled ascent →\n• Diver fatality / serious injury\n• DCI requiring hyperbaric treatment\n• Umbilical damage\n• Emergency decompression required\n• Full ERP activation',
                'controls': '1. Ascent rate monitored — max 30 fsw/min (USN Rev7).\n2. Decompression stops confirmed by diver.\n3. Surface interval monitoring per dive tables.\n4. IMCA D078 — active umbilical tending during ascent.\n5. Diver reports to tender before ascent.\n6. Decompression chamber on 18-hour bend watch.\n7. Standby diver ready for emergency deployment.\n8. Diver dressed in during ascent including in-water stops.',
                'recovery': 'Stop ascent.\nRe-descend to missed stop.\nResume decompression.\nDMT assessment.\nDDC ready for treatment.\nERP if serious DCI.',
                'action': 'Diving Supervisor / Tender / Diver / Diver Medic / Standby Diver',
                'p': '3C', 's': '4C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '5.2',
                'task': 'Post-dive equipment inspection and log completion',
                'hazard': '• High pressure systems\n• Hot equipment\n• Gas cylinder handling',
                'threat': 'Overpressurised cylinder.\nCompressor hot surfaces.\nGas cylinder dropped.\nContaminated air from faulty compressor.',
                'consequence': 'Equipment damage or injury →\n• Cylinder rupture\n• Burns from compressor\n• Crush injury\n• Contaminated air for next dive',
                'controls': '1. All cylinders pressure tested per IMCA D043.\n2. Compressor serviced per PMS.\n3. Air analysis before next dive — BS EN 12012.\n4. IMCA D018 — examination of diving plant.\n5. Post-dive equipment checklist completed.\n6. Dive log completed with all details.\n7. Umbilical inspected for damage.',
                'recovery': 'Isolate contaminated system.\nFlush and retest.\nReplace compromised cylinders.\nDMT assessment.\nNotify Diving Superintendent.',
                'action': 'Dive Technician / Diving Supervisor / Diver Medic',
                'p': '2C', 's': '2C', 'rp': '0', 'rs': '0',
            },
            {
                'no': '5.3',
                'task': 'Project close-out — documentation and handover',
                'hazard': '• Incomplete documentation\n• Missing records\n• Client sign-off delay',
                'threat': 'Dive log not completed.\nPTW not closed out.\nIsolation certificates not returned.\nPhoto documentation incomplete.',
                'consequence': 'Administrative failure →\n• Regulatory non-compliance\n• Client拒签 — payment delay\n• Audit failure\n• Legal liability',
                'controls': '1. All dive logs completed and signed.\n2. PTW closed out with all signatures.\n3. LOTO certificates returned and filed.\n4. Photo/video documentation archived.\n5. HIRA workbook delivered to client.\n6. Attendance sheet signed by all attendees.\n7. Post-project debrief with client.',
                'recovery': 'Complete missing documentation.\nObtain client sign-off.\nFile all records.\nSchedule debrief meeting.',
                'action': 'Diving Supervisor / Project Manager / HSE Lead / Client Rep / PIC',
                'p': '2C', 's': '2C', 'rp': '0', 'rs': '0',
            },
        ]
    })

    wb = create_hira_workbook(
        client=client, contractor=contractor, project=project,
        vessel=vessel, location=location, worksite=worksite,
        doc_ref=doc_ref, date=date, phases=phases,
        title_short="SAL Buoy"
    )
    return wb


# ==============================================================================
# MAIN — Generate all 4 workbooks
# ==============================================================================
if __name__ == '__main__':
    today = datetime.now().strftime('%Y-%m-%d')

    projects = [
        ("Snake Island", gen_project_1, "Snake Island Wrecks Removal"),
        ("SNEPCo Bonga", gen_project_2, "SNEPCo Bonga IWS Mooring"),
        ("3D Photogrammetry", gen_project_3, "3D Photogrammetry Inspection"),
        ("SAL Flange", gen_project_4, "SAL Flange Disconnection/Reconnection"),
    ]

    results = []
    for folder_name, gen_func, display_name in projects:
        print(f"\nGenerating HIRA for: {display_name}")
        try:
            wb = gen_func()
            out_path = os.path.join(SAVE_DIR, f"{today}-{folder_name}_HIRA.xlsx")
            wb.save(out_path)
            size = os.path.getsize(out_path)
            print(f"  Saved: {out_path} ({size:,} bytes)")
            results.append({
                'folder_name': folder_name,
                'display_name': display_name,
                'path': out_path,
                'size': size,
                'status': 'SUCCESS',
            })
        except Exception as e:
            print(f"  ERROR: {e}")
            import traceback
            traceback.print_exc()
            results.append({
                'folder_name': folder_name,
                'display_name': display_name,
                'path': '',
                'size': 0,
                'status': f'ERROR: {e}',
            })

    print(f"\n{'='*60}")
    print(f"Generation complete: {sum(1 for r in results if r['status'] == 'SUCCESS')}/{len(results)} successful")
    for r in results:
        print(f"  {r['display_name']}: {r['status']}")
