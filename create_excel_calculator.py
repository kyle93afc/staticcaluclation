import pandas as pd
import numpy as np
from math import pi, sqrt
import xlsxwriter
from datetime import datetime

def create_chimney_calculator_excel():
    """Create a comprehensive Excel chimney calculator with validation"""

    # Create Excel workbook
    workbook = xlsxwriter.Workbook('Steel_Chimney_Calculator.xlsx')

    # Define formats
    formats = {
        'header': workbook.add_format({
            'bold': True, 'font_size': 14, 'bg_color': '#2F5597',
            'font_color': 'white', 'align': 'center', 'border': 1
        }),
        'input': workbook.add_format({
            'bg_color': '#E6F3FF', 'border': 1, 'num_format': '0.00'
        }),
        'calc': workbook.add_format({
            'bg_color': '#F0F0F0', 'border': 1, 'num_format': '0.00'
        }),
        'result': workbook.add_format({
            'bg_color': '#D4F6D4', 'border': 1, 'num_format': '0.00', 'bold': True
        }),
        'warning': workbook.add_format({
            'bg_color': '#FFE6CC', 'border': 1, 'font_color': '#D2691E'
        }),
        'error': workbook.add_format({
            'bg_color': '#FFE6E6', 'border': 1, 'font_color': 'red', 'bold': True
        }),
        'ok': workbook.add_format({
            'bg_color': '#E6FFE6', 'border': 1, 'font_color': 'green', 'bold': True
        }),
        'title': workbook.add_format({
            'bold': True, 'font_size': 16, 'bg_color': '#1F4E79',
            'font_color': 'white', 'align': 'center'
        })
    }

    # Create main calculation worksheet
    ws_main = workbook.add_worksheet('Chimney Calculator')
    ws_main.set_column('A:A', 25)
    ws_main.set_column('B:B', 15)
    ws_main.set_column('C:C', 8)
    ws_main.set_column('D:D', 25)
    ws_main.set_column('E:E', 15)
    ws_main.set_column('F:F', 15)

    # Title
    ws_main.merge_range('A1:F1', 'STEEL CHIMNEY STATIC CALCULATION', formats['title'])
    ws_main.merge_range('A2:F2', 'According to EN 1993-3-2, EN 1991-1-4', formats['header'])

    row = 4

    # INPUT PARAMETERS SECTION
    ws_main.write(row, 0, 'INPUT PARAMETERS', formats['header'])
    ws_main.write(row, 1, 'Value', formats['header'])
    ws_main.write(row, 2, 'Unit', formats['header'])
    ws_main.write(row, 3, 'Status Check', formats['header'])
    ws_main.write(row, 4, 'Min Value', formats['header'])
    ws_main.write(row, 5, 'Max Value', formats['header'])
    row += 1

    # Main Dimensions
    ws_main.write(row, 0, 'MAIN DIMENSIONS')
    row += 1

    # Height
    ws_main.write(row, 0, 'Chimney Height (H)')
    ws_main.write(row, 1, 13.50, formats['input'])
    ws_main.write(row, 2, 'm')
    ws_main.write(row, 3, '=IF(AND(B' + str(row+1) + '>=E' + str(row+1) + ',B' + str(row+1) + '<=F' + str(row+1) + '),"✓ OK","⚠ Check Range")')
    ws_main.write(row, 4, 5.0)
    ws_main.write(row, 5, 100.0)
    height_cell = f'B{row+1}'
    row += 1

    # Diameter
    ws_main.write(row, 0, 'Outside Diameter (D)')
    ws_main.write(row, 1, 1422, formats['input'])
    ws_main.write(row, 2, 'mm')
    ws_main.write(row, 3, '=IF(AND(B' + str(row+1) + '>=E' + str(row+1) + ',B' + str(row+1) + '<=F' + str(row+1) + '),"✓ OK","⚠ Check Range")')
    ws_main.write(row, 4, 500)
    ws_main.write(row, 5, 5000)
    diameter_cell = f'B{row+1}'
    row += 1

    # Shell thickness
    ws_main.write(row, 0, 'Shell Thickness (t)')
    ws_main.write(row, 1, 8.0, formats['input'])
    ws_main.write(row, 2, 'mm')
    ws_main.write(row, 3, '=IF(AND(B' + str(row+1) + '>=E' + str(row+1) + ',B' + str(row+1) + '<=F' + str(row+1) + '),"✓ OK","⚠ Check Range")')
    ws_main.write(row, 4, 3.0)
    ws_main.write(row, 5, 25.0)
    thickness_cell = f'B{row+1}'
    row += 1

    # Corrosion allowance
    ws_main.write(row, 0, 'Corrosion Allowance')
    ws_main.write(row, 1, 0.35, formats['input'])
    ws_main.write(row, 2, 'mm')
    ws_main.write(row, 3, '=IF(B' + str(row+1) + '<B' + str(row) + ',"✓ OK","❌ > Shell Thickness")')
    corr_cell = f'B{row+1}'
    row += 2

    # Wind Parameters
    ws_main.write(row, 0, 'WIND PARAMETERS')
    row += 1

    # Basic wind velocity
    ws_main.write(row, 0, 'Basic Wind Velocity (vb)')
    ws_main.write(row, 1, 25.5, formats['input'])
    ws_main.write(row, 2, 'm/s')
    ws_main.write(row, 3, '=IF(AND(B' + str(row+1) + '>=E' + str(row+1) + ',B' + str(row+1) + '<=F' + str(row+1) + '),"✓ OK","⚠ Check Range")')
    ws_main.write(row, 4, 15.0)
    ws_main.write(row, 5, 50.0)
    wind_vel_cell = f'B{row+1}'
    row += 1

    # Site altitude
    ws_main.write(row, 0, 'Site Altitude')
    ws_main.write(row, 1, 200, formats['input'])
    ws_main.write(row, 2, 'm')
    ws_main.write(row, 3, '=IF(B' + str(row+1) + '>=0,"✓ OK","❌ Negative Value")')
    altitude_cell = f'B{row+1}'
    row += 2

    # Material Properties
    ws_main.write(row, 0, 'MATERIAL PROPERTIES')
    row += 1

    # Steel grade
    ws_main.write(row, 0, 'Steel Grade')
    ws_main.write(row, 1, 'S235JR', formats['input'])
    ws_main.write(row, 2, '-')
    ws_main.write(row, 3, '✓ Standard Grade')
    row += 1

    # Yield strength
    ws_main.write(row, 0, 'Yield Strength (fy)')
    ws_main.write(row, 1, 235, formats['input'])
    ws_main.write(row, 2, 'N/mm²')
    ws_main.write(row, 3, '=IF(B' + str(row+1) + '>0,"✓ OK","❌ Invalid")')
    fy_cell = f'B{row+1}'
    row += 3

    # CALCULATION SECTION
    calc_start_row = row
    ws_main.write(row, 0, 'CALCULATIONS', formats['header'])
    ws_main.write(row, 1, 'Result', formats['header'])
    ws_main.write(row, 2, 'Unit', formats['header'])
    ws_main.write(row, 3, 'Formula/Check', formats['header'])
    ws_main.write(row, 4, 'Status', formats['header'])
    row += 1

    # Basic calculations
    ws_main.write(row, 0, 'Altitude Factor (c_alt)')
    ws_main.write(row, 1, f'=1+0.001*{altitude_cell}', formats['calc'])
    ws_main.write(row, 2, '-')
    ws_main.write(row, 3, '1 + 0.001 × altitude')
    ws_main.write(row, 4, '=IF(B' + str(row+1) + '>0.95,"✓ OK","⚠ Low")')
    c_alt_cell = f'B{row+1}'
    row += 1

    # Basic wind velocity
    ws_main.write(row, 0, 'Basic Wind Velocity (vb0)')
    ws_main.write(row, 1, f'={wind_vel_cell}*{c_alt_cell}', formats['calc'])
    ws_main.write(row, 2, 'm/s')
    ws_main.write(row, 3, 'vb_map × c_alt')
    ws_main.write(row, 4, '=IF(B' + str(row+1) + '<50,"✓ OK","⚠ High Wind")')
    vb0_cell = f'B{row+1}'
    row += 1

    # Basic velocity pressure
    ws_main.write(row, 0, 'Basic Velocity Pressure (qb)')
    ws_main.write(row, 1, f'=0.5*1.226*{vb0_cell}^2/1000', formats['calc'])
    ws_main.write(row, 2, 'kN/m²')
    ws_main.write(row, 3, '0.5 × ρ × vb0²')
    qb_cell = f'B{row+1}'
    row += 1

    # Peak velocity pressure (simplified)
    ws_main.write(row, 0, 'Peak Velocity Pressure (qp)')
    ws_main.write(row, 1, f'=2.36*{qb_cell}', formats['calc'])
    ws_main.write(row, 2, 'kN/m²')
    ws_main.write(row, 3, 'Ce × qb (simplified)')
    qp_cell = f'B{row+1}'
    row += 1

    # Cross-sectional area
    ws_main.write(row, 0, 'Cross-sectional Area (A)')
    ws_main.write(row, 1, f'=PI()*{diameter_cell}*({thickness_cell}-{corr_cell})', formats['calc'])
    ws_main.write(row, 2, 'mm²')
    ws_main.write(row, 3, 'π × D × t_corroded')
    area_cell = f'B{row+1}'
    row += 1

    # Section modulus
    ws_main.write(row, 0, 'Section Modulus (W)')
    ws_main.write(row, 1, f'=PI()*({diameter_cell}-({thickness_cell}-{corr_cell}))^2*({thickness_cell}-{corr_cell})/4', formats['calc'])
    ws_main.write(row, 2, 'mm³')
    ws_main.write(row, 3, 'π × Dm² × t / 4')
    section_mod_cell = f'B{row+1}'
    row += 1

    # Wind load per meter
    ws_main.write(row, 0, 'Wind Load per Meter (fw)')
    ws_main.write(row, 1, f'=1.4*0.965*{qp_cell}*{diameter_cell}/1000*0.556', formats['calc'])
    ws_main.write(row, 2, 'kN/m')
    ws_main.write(row, 3, 'γQ × CsCd × qp × D × Cf')
    fw_cell = f'B{row+1}'
    row += 2

    # RESULTS SECTION
    ws_main.write(row, 0, 'KEY RESULTS', formats['header'])
    ws_main.write(row, 1, 'Value', formats['header'])
    ws_main.write(row, 2, 'Unit', formats['header'])
    ws_main.write(row, 3, 'Check', formats['header'])
    ws_main.write(row, 4, 'Status', formats['header'])
    row += 1

    # Foundation loads
    ws_main.write(row, 0, 'Total Wind Force (Q)')
    ws_main.write(row, 1, f'={fw_cell}*{height_cell}', formats['result'])
    ws_main.write(row, 2, 'kN')
    ws_main.write(row, 3, 'fw × H')
    total_force_cell = f'B{row+1}'
    row += 1

    ws_main.write(row, 0, 'Base Moment (M)')
    ws_main.write(row, 1, f'={fw_cell}*{height_cell}^2/2', formats['result'])
    ws_main.write(row, 2, 'kNm')
    ws_main.write(row, 3, 'fw × H² / 2')
    base_moment_cell = f'B{row+1}'
    row += 1

    # Maximum stress
    ws_main.write(row, 0, 'Maximum Stress (σmax)')
    ws_main.write(row, 1, f'={base_moment_cell}*1000000/{section_mod_cell}', formats['result'])
    ws_main.write(row, 2, 'N/mm²')
    ws_main.write(row, 3, 'M / W')
    max_stress_cell = f'B{row+1}'

    # Stress check
    ws_main.write(row, 4, f'=IF({max_stress_cell}<0.8*{fy_cell},"✓ OK - Low","IF({max_stress_cell}<{fy_cell},"⚠ OK - High","❌ OVERSTRESSED"))')
    row += 1

    # Stress utilization
    ws_main.write(row, 0, 'Stress Utilization')
    ws_main.write(row, 1, f'={max_stress_cell}/{fy_cell}', formats['result'])
    ws_main.write(row, 2, '-')
    ws_main.write(row, 3, 'σmax / fy')
    util_cell = f'B{row+1}'
    ws_main.write(row, 4, f'=IF({util_cell}<0.8,"✓ SAFE","IF({util_cell}<1,"⚠ CHECK","❌ FAIL"))')
    row += 2

    # DESIGN STATUS
    ws_main.write(row, 0, 'OVERALL DESIGN STATUS', formats['header'])
    row += 1

    # Overall status check
    ws_main.write(row, 0, 'Design Verdict')
    ws_main.write(row, 1, f'=IF({util_cell}<0.8,"✓ DESIGN OK - SAFE","IF({util_cell}<1,"⚠ DESIGN OK - HIGH UTILIZATION","❌ DESIGN FAILS - OVERSTRESSED"))')
    ws_main.write(row, 3, 'Based on stress utilization')

    # Apply conditional formatting for status
    ws_main.conditional_format(f'B{row+1}:B{row+1}', {
        'type': 'text',
        'criteria': 'containing',
        'value': '✓',
        'format': formats['ok']
    })

    ws_main.conditional_format(f'B{row+1}:B{row+1}', {
        'type': 'text',
        'criteria': 'containing',
        'value': '⚠',
        'format': formats['warning']
    })

    ws_main.conditional_format(f'B{row+1}:B{row+1}', {
        'type': 'text',
        'criteria': 'containing',
        'value': '❌',
        'format': formats['error']
    })

    row += 3

    # DESIGN RECOMMENDATIONS
    ws_main.write(row, 0, 'DESIGN RECOMMENDATIONS', formats['header'])
    row += 1

    ws_main.write(row, 0, 'If Overstressed:')
    row += 1
    ws_main.write(row, 0, '• Increase shell thickness')
    row += 1
    ws_main.write(row, 0, '• Increase diameter')
    row += 1
    ws_main.write(row, 0, '• Use higher grade steel')
    row += 1
    ws_main.write(row, 0, '• Add stiffening rings')
    row += 2

    ws_main.write(row, 0, 'If High Utilization (>80%):')
    row += 1
    ws_main.write(row, 0, '• Consider slight thickness increase')
    row += 1
    ws_main.write(row, 0, '• Verify detailed calculations')
    row += 1
    ws_main.write(row, 0, '• Check buckling requirements')

    # Create Load Table worksheet
    create_load_table_worksheet(workbook, formats)

    # Create Instructions worksheet
    create_instructions_worksheet(workbook, formats)

    # Create Validation worksheet
    create_validation_worksheet(workbook, formats)

    workbook.close()
    print("Excel calculator created: Steel_Chimney_Calculator.xlsx")

def create_load_table_worksheet(workbook, formats):
    """Create detailed load distribution table"""
    ws = workbook.add_worksheet('Load Distribution')
    ws.set_column('A:A', 12)
    ws.set_column('B:G', 10)

    # Title
    ws.merge_range('A1:G1', 'WIND LOAD DISTRIBUTION', formats['title'])

    # Headers
    headers = ['Height', 'From', 'To', 'Wind Load', 'Moment Arm', 'Moment', 'Cumulative']
    units = ['[m]', '[mm]', '[mm]', '[kN/m]', '[m]', '[kNm]', '[kNm]']

    for i, (header, unit) in enumerate(zip(headers, units)):
        ws.write(2, i, header, formats['header'])
        ws.write(3, i, unit, formats['header'])

    # Sample data (would be calculated from main sheet)
    segments = [
        (0.075, 0, 150, '=Calculator!B25', 0.075, '=D5*E5', '=F5'),
        (0.250, 150, 350, '=Calculator!B25', 0.250, '=D6*E6', '=F6+G5'),
        (0.425, 350, 500, '=Calculator!B25', 0.425, '=D7*E7', '=F7+G6'),
    ]

    for i, (height, from_h, to_h, load_ref, arm, moment_formula, cum_formula) in enumerate(segments):
        row = 4 + i
        ws.write(row, 0, height, formats['calc'])
        ws.write(row, 1, from_h, formats['calc'])
        ws.write(row, 2, to_h, formats['calc'])
        ws.write(row, 3, load_ref if isinstance(load_ref, str) else load_ref, formats['calc'])
        ws.write(row, 4, arm, formats['calc'])
        ws.write(row, 5, moment_formula, formats['calc'])
        ws.write(row, 6, cum_formula, formats['result'])

def create_instructions_worksheet(workbook, formats):
    """Create user instructions"""
    ws = workbook.add_worksheet('Instructions')
    ws.set_column('A:A', 80)

    instructions = [
        ("🏭 STEEL CHIMNEY CALCULATOR - USER GUIDE", formats['title']),
        ("", None),
        ("HOW TO USE:", formats['header']),
        ("", None),
        ("1. INPUT PARAMETERS (Blue cells):", formats['header']),
        ("   • Modify only the BLUE cells in the 'Chimney Calculator' sheet", None),
        ("   • Enter your chimney dimensions, wind data, and material properties", None),
        ("   • Values will be validated automatically", None),
        ("", None),
        ("2. CHECK STATUS INDICATORS:", formats['header']),
        ("   ✓ OK = Value is acceptable", None),
        ("   ⚠ Warning = Check the value, may be at limits", None),
        ("   ❌ Error = Value is outside acceptable range", None),
        ("", None),
        ("3. REVIEW RESULTS:", formats['header']),
        ("   • Green cells show key results", None),
        ("   • Check the 'Overall Design Status' at the bottom", None),
        ("   • Review stress utilization ratio", None),
        ("", None),
        ("4. DESIGN DECISIONS:", formats['header']),
        ("   • Stress utilization < 80% = Safe design", None),
        ("   • Stress utilization 80-100% = OK but high", None),
        ("   • Stress utilization > 100% = Overstressed, redesign needed", None),
        ("", None),
        ("EXAMPLE MODIFICATIONS:", formats['header']),
        ("", None),
        ("For Taller Chimney:", None),
        ("   • Increase Height (H)", None),
        ("   • May need to increase Diameter or Thickness", None),
        ("", None),
        ("For Higher Wind Zone:", None),
        ("   • Increase Basic Wind Velocity", None),
        ("   • Check if thickness needs increasing", None),
        ("", None),
        ("For Stronger Design:", None),
        ("   • Increase Shell Thickness", None),
        ("   • Use higher grade steel (S355 instead of S235)", None),
        ("", None),
        ("⚠️ IMPORTANT NOTES:", formats['warning']),
        ("", None),
        ("• This is a SIMPLIFIED calculation for preliminary design", None),
        ("• Final design must include:", None),
        ("  - Detailed buckling checks", None),
        ("  - Natural frequency analysis", None),
        ("  - Foundation design", None),
        ("  - Local reinforcement at openings", None),
        ("  - Fatigue analysis if applicable", None),
        ("", None),
        ("• Always verify with commercial software", None),
        ("• Have design checked by qualified engineer", None),
        ("• Comply with local building codes", None),
    ]

    for i, (text, fmt) in enumerate(instructions):
        if fmt:
            ws.write(i, 0, text, fmt)
        else:
            ws.write(i, 0, text)

def create_validation_worksheet(workbook, formats):
    """Create validation and checking worksheet"""
    ws = workbook.add_worksheet('Validation')
    ws.set_column('A:A', 30)
    ws.set_column('B:D', 15)

    ws.write(0, 0, 'VALIDATION CHECKS', formats['title'])

    checks = [
        ("Parameter", "Current Value", "Valid Range", "Status"),
        ("Slenderness Ratio", "=Calculator!B6/Calculator!B7*1000", "5 - 50", "=IF(B3<50,\"✓ OK\",\"⚠ High\")"),
        ("Thickness Ratio", "=Calculator!B7/Calculator!B6", "0.002 - 0.020", "=IF(AND(B4>=0.002,B4<=0.020),\"✓ OK\",\"⚠ Check\")"),
        ("Wind Pressure", "=Calculator!B20", "< 2.0 kN/m²", "=IF(B5<2,\"✓ OK\",\"⚠ High Wind\")"),
        ("Stress Level", "=Calculator!B32", "< 235 N/mm²", "=IF(B6<235,\"✓ OK\",\"❌ Overstressed\")"),
    ]

    for i, check in enumerate(checks):
        for j, value in enumerate(check):
            if i == 0:
                ws.write(i+2, j, value, formats['header'])
            else:
                if j == 1 or j == 3:  # Formula cells
                    ws.write(i+2, j, value, formats['calc'])
                else:
                    ws.write(i+2, j, value)

if __name__ == "__main__":
    create_chimney_calculator_excel()