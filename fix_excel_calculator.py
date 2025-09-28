import pandas as pd
import numpy as np
from math import pi, sqrt
import xlsxwriter
from datetime import datetime

def create_fixed_chimney_calculator():
    """Create Excel calculator with proper internal references only"""

    # Create Excel workbook
    workbook = xlsxwriter.Workbook('Steel_Chimney_Calculator_Fixed.xlsx')

    # Define formats
    formats = {
        'header': workbook.add_format({
            'bold': True, 'font_size': 14, 'bg_color': '#2F5597',
            'font_color': 'white', 'align': 'center', 'border': 1
        }),
        'input': workbook.add_format({
            'bg_color': '#E6F3FF', 'border': 1, 'num_format': '0.00', 'bold': True
        }),
        'calc': workbook.add_format({
            'bg_color': '#F0F0F0', 'border': 1, 'num_format': '0.00'
        }),
        'result': workbook.add_format({
            'bg_color': '#D4F6D4', 'border': 1, 'num_format': '0.00', 'bold': True
        }),
        'warning': workbook.add_format({
            'bg_color': '#FFE6CC', 'border': 1, 'font_color': '#D2691E', 'bold': True
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
        }),
        'text': workbook.add_format({
            'border': 1, 'text_wrap': True
        })
    }

    # Create main calculation worksheet
    ws = workbook.add_worksheet('Calculator')
    ws.set_column('A:A', 30)
    ws.set_column('B:B', 15)
    ws.set_column('C:C', 8)
    ws.set_column('D:D', 35)
    ws.set_column('E:E', 15)
    ws.set_column('F:F', 15)

    # Title
    ws.merge_range('A1:F1', 'STEEL CHIMNEY STATIC CALCULATION', formats['title'])
    ws.merge_range('A2:F2', 'According to EN 1993-3-2, EN 1991-1-4', formats['header'])

    row = 4

    # INPUT PARAMETERS SECTION
    ws.write(row, 0, 'INPUT PARAMETERS', formats['header'])
    ws.write(row, 1, 'Value', formats['header'])
    ws.write(row, 2, 'Unit', formats['header'])
    ws.write(row, 3, 'Status Check', formats['header'])
    ws.write(row, 4, 'Min Value', formats['header'])
    ws.write(row, 5, 'Max Value', formats['header'])
    row += 1

    # Main Dimensions
    ws.write(row, 0, 'MAIN DIMENSIONS', formats['header'])
    row += 1

    # Height
    ws.write(row, 0, 'Chimney Height (H)')
    ws.write(row, 1, 13.50, formats['input'])
    ws.write(row, 2, 'm')
    ws.write(row, 3, f'=IF(AND(B{row+1}>=E{row+1},B{row+1}<=F{row+1}),"✓ OK","⚠ Check Range")', formats['text'])
    ws.write(row, 4, 5.0)
    ws.write(row, 5, 100.0)
    height_row = row + 1
    row += 1

    # Diameter
    ws.write(row, 0, 'Outside Diameter (D)')
    ws.write(row, 1, 1422, formats['input'])
    ws.write(row, 2, 'mm')
    ws.write(row, 3, f'=IF(AND(B{row+1}>=E{row+1},B{row+1}<=F{row+1}),"✓ OK","⚠ Check Range")', formats['text'])
    ws.write(row, 4, 500)
    ws.write(row, 5, 5000)
    diameter_row = row + 1
    row += 1

    # Shell thickness
    ws.write(row, 0, 'Shell Thickness (t)')
    ws.write(row, 1, 8.0, formats['input'])
    ws.write(row, 2, 'mm')
    ws.write(row, 3, f'=IF(AND(B{row+1}>=E{row+1},B{row+1}<=F{row+1}),"✓ OK","⚠ Check Range")', formats['text'])
    ws.write(row, 4, 3.0)
    ws.write(row, 5, 25.0)
    thickness_row = row + 1
    row += 1

    # Corrosion allowance
    ws.write(row, 0, 'Corrosion Allowance')
    ws.write(row, 1, 0.35, formats['input'])
    ws.write(row, 2, 'mm')
    ws.write(row, 3, f'=IF(B{row+1}<B{thickness_row},"✓ OK","❌ > Shell Thickness")', formats['text'])
    corr_row = row + 1
    row += 2

    # Wind Parameters
    ws.write(row, 0, 'WIND PARAMETERS', formats['header'])
    row += 1

    # Basic wind velocity
    ws.write(row, 0, 'Basic Wind Velocity (vb)')
    ws.write(row, 1, 25.5, formats['input'])
    ws.write(row, 2, 'm/s')
    ws.write(row, 3, f'=IF(AND(B{row+1}>=E{row+1},B{row+1}<=F{row+1}),"✓ OK","⚠ Check Range")', formats['text'])
    ws.write(row, 4, 15.0)
    ws.write(row, 5, 50.0)
    wind_vel_row = row + 1
    row += 1

    # Site altitude
    ws.write(row, 0, 'Site Altitude')
    ws.write(row, 1, 200, formats['input'])
    ws.write(row, 2, 'm')
    ws.write(row, 3, f'=IF(B{row+1}>=0,"✓ OK","❌ Negative Value")', formats['text'])
    altitude_row = row + 1
    row += 2

    # Material Properties
    ws.write(row, 0, 'MATERIAL PROPERTIES', formats['header'])
    row += 1

    # Steel grade
    ws.write(row, 0, 'Steel Grade')
    ws.write(row, 1, 'S235JR', formats['input'])
    ws.write(row, 2, '-')
    ws.write(row, 3, '✓ Standard Grade', formats['text'])
    row += 1

    # Yield strength
    ws.write(row, 0, 'Yield Strength (fy)')
    ws.write(row, 1, 235, formats['input'])
    ws.write(row, 2, 'N/mm²')
    ws.write(row, 3, f'=IF(B{row+1}>0,"✓ OK","❌ Invalid")', formats['text'])
    fy_row = row + 1
    row += 3

    # CALCULATION SECTION
    ws.write(row, 0, 'CALCULATIONS', formats['header'])
    ws.write(row, 1, 'Result', formats['header'])
    ws.write(row, 2, 'Unit', formats['header'])
    ws.write(row, 3, 'Formula/Check', formats['header'])
    ws.write(row, 4, 'Status', formats['header'])
    row += 1

    # Basic calculations
    ws.write(row, 0, 'Altitude Factor (c_alt)')
    ws.write(row, 1, f'=1+0.001*B{altitude_row}', formats['calc'])
    ws.write(row, 2, '-')
    ws.write(row, 3, '1 + 0.001 × altitude')
    ws.write(row, 4, f'=IF(B{row+1}>0.95,"✓ OK","⚠ Low")', formats['text'])
    c_alt_row = row + 1
    row += 1

    # Basic wind velocity
    ws.write(row, 0, 'Basic Wind Velocity (vb0)')
    ws.write(row, 1, f'=B{wind_vel_row}*B{c_alt_row}', formats['calc'])
    ws.write(row, 2, 'm/s')
    ws.write(row, 3, 'vb_map × c_alt')
    ws.write(row, 4, f'=IF(B{row+1}<50,"✓ OK","⚠ High Wind")', formats['text'])
    vb0_row = row + 1
    row += 1

    # Basic velocity pressure
    ws.write(row, 0, 'Basic Velocity Pressure (qb)')
    ws.write(row, 1, f'=0.5*1.226*B{vb0_row}^2/1000', formats['calc'])
    ws.write(row, 2, 'kN/m²')
    ws.write(row, 3, '0.5 × ρ × vb0²')
    qb_row = row + 1
    row += 1

    # Peak velocity pressure (simplified)
    ws.write(row, 0, 'Peak Velocity Pressure (qp)')
    ws.write(row, 1, f'=2.36*B{qb_row}', formats['calc'])
    ws.write(row, 2, 'kN/m²')
    ws.write(row, 3, 'Ce × qb (simplified)')
    qp_row = row + 1
    row += 1

    # Cross-sectional area
    ws.write(row, 0, 'Cross-sectional Area (A)')
    ws.write(row, 1, f'=PI()*B{diameter_row}*(B{thickness_row}-B{corr_row})', formats['calc'])
    ws.write(row, 2, 'mm²')
    ws.write(row, 3, 'π × D × t_corroded')
    area_row = row + 1
    row += 1

    # Section modulus
    ws.write(row, 0, 'Section Modulus (W)')
    ws.write(row, 1, f'=PI()*(B{diameter_row}-(B{thickness_row}-B{corr_row}))^2*(B{thickness_row}-B{corr_row})/4', formats['calc'])
    ws.write(row, 2, 'mm³')
    ws.write(row, 3, 'π × Dm² × t / 4')
    section_mod_row = row + 1
    row += 1

    # Wind load per meter
    ws.write(row, 0, 'Wind Load per Meter (fw)')
    ws.write(row, 1, f'=1.4*0.965*B{qp_row}*B{diameter_row}/1000*0.556', formats['calc'])
    ws.write(row, 2, 'kN/m')
    ws.write(row, 3, 'γQ × CsCd × qp × D × Cf')
    fw_row = row + 1
    row += 2

    # RESULTS SECTION
    ws.write(row, 0, 'KEY RESULTS', formats['header'])
    ws.write(row, 1, 'Value', formats['header'])
    ws.write(row, 2, 'Unit', formats['header'])
    ws.write(row, 3, 'Check', formats['header'])
    ws.write(row, 4, 'Status', formats['header'])
    row += 1

    # Foundation loads
    ws.write(row, 0, 'Total Wind Force (Q)')
    ws.write(row, 1, f'=B{fw_row}*B{height_row}', formats['result'])
    ws.write(row, 2, 'kN')
    ws.write(row, 3, 'fw × H')
    total_force_row = row + 1
    row += 1

    ws.write(row, 0, 'Base Moment (M)')
    ws.write(row, 1, f'=B{fw_row}*B{height_row}^2/2', formats['result'])
    ws.write(row, 2, 'kNm')
    ws.write(row, 3, 'fw × H² / 2')
    base_moment_row = row + 1
    row += 1

    # Maximum stress
    ws.write(row, 0, 'Maximum Stress (σmax)')
    ws.write(row, 1, f'=B{base_moment_row}*1000000/B{section_mod_row}', formats['result'])
    ws.write(row, 2, 'N/mm²')
    ws.write(row, 3, 'M / W')
    max_stress_row = row + 1

    # Stress check
    ws.write(row, 4, f'=IF(B{max_stress_row}<0.8*B{fy_row},"✓ OK - Low","IF(B{max_stress_row}<B{fy_row},"⚠ OK - High","❌ OVERSTRESSED"))', formats['text'])
    row += 1

    # Stress utilization
    ws.write(row, 0, 'Stress Utilization')
    ws.write(row, 1, f'=B{max_stress_row}/B{fy_row}', formats['result'])
    ws.write(row, 2, '-')
    ws.write(row, 3, 'σmax / fy')
    util_row = row + 1
    ws.write(row, 4, f'=IF(B{util_row}<0.8,"✓ SAFE","IF(B{util_row}<1,"⚠ CHECK","❌ FAIL"))', formats['text'])
    row += 2

    # DESIGN STATUS
    ws.write(row, 0, 'OVERALL DESIGN STATUS', formats['header'])
    row += 1

    # Overall status check
    ws.write(row, 0, 'Design Verdict')
    ws.merge_range(f'B{row+1}:D{row+1}', f'=IF(B{util_row}<0.8,"✓ DESIGN OK - SAFE","IF(B{util_row}<1,"⚠ DESIGN OK - HIGH UTILIZATION","❌ DESIGN FAILS - OVERSTRESSED"))')
    ws.write(row, 4, 'Based on stress utilization')
    row += 3

    # DESIGN RECOMMENDATIONS
    ws.write(row, 0, 'DESIGN RECOMMENDATIONS', formats['header'])
    row += 1

    recommendations = [
        'If Overstressed:',
        '• Increase shell thickness',
        '• Increase diameter',
        '• Use higher grade steel (S355 instead of S235)',
        '• Add stiffening rings',
        '',
        'If High Utilization (>80%):',
        '• Consider slight thickness increase',
        '• Verify detailed calculations',
        '• Check buckling requirements',
        '• Use commercial software for verification',
        '',
        'Always Required:',
        '• Buckling analysis per EN 1993-1-6',
        '• Natural frequency check',
        '• Foundation design',
        '• Local reinforcement at openings'
    ]

    for rec in recommendations:
        ws.write(row, 0, rec)
        row += 1

    # Add conditional formatting for the design verdict cell
    verdict_range = f'B{row-len(recommendations)-2}:D{row-len(recommendations)-2}'

    # Format for OK status
    ws.conditional_format(verdict_range, {
        'type': 'text',
        'criteria': 'containing',
        'value': '✓',
        'format': formats['ok']
    })

    # Format for warning status
    ws.conditional_format(verdict_range, {
        'type': 'text',
        'criteria': 'containing',
        'value': '⚠',
        'format': formats['warning']
    })

    # Format for error status
    ws.conditional_format(verdict_range, {
        'type': 'text',
        'criteria': 'containing',
        'value': '❌',
        'format': formats['error']
    })

    # Create Instructions worksheet
    create_instructions_worksheet(workbook, formats)

    # Create Validation worksheet
    create_validation_worksheet(workbook, formats, height_row, diameter_row, thickness_row, wind_vel_row, max_stress_row, fy_row, util_row)

    workbook.close()
    print("Fixed Excel calculator created: Steel_Chimney_Calculator_Fixed.xlsx")

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
        ("   • Modify only the BLUE cells in the 'Calculator' sheet", None),
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
        ("   • Check the 'Overall Design Status' section", None),
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

def create_validation_worksheet(workbook, formats, height_row, diameter_row, thickness_row, wind_vel_row, max_stress_row, fy_row, util_row):
    """Create validation and checking worksheet"""
    ws = workbook.add_worksheet('Validation')
    ws.set_column('A:A', 30)
    ws.set_column('B:D', 15)

    ws.write(0, 0, 'VALIDATION CHECKS', formats['title'])

    checks = [
        ("Parameter", "Current Value", "Valid Range", "Status"),
        ("Slenderness Ratio", f"=Calculator!B{height_row}/Calculator!B{diameter_row}*1000", "5 - 50", "=IF(B3<50,\"✓ OK\",\"⚠ High\")"),
        ("Thickness Ratio", f"=Calculator!B{thickness_row}/Calculator!B{diameter_row}", "0.002 - 0.020", "=IF(AND(B4>=0.002,B4<=0.020),\"✓ OK\",\"⚠ Check\")"),
        ("Wind Velocity", f"=Calculator!B{wind_vel_row}", "< 35 m/s", "=IF(B5<35,\"✓ OK\",\"⚠ High Wind\")"),
        ("Stress Level", f"=Calculator!B{max_stress_row}", f"< Calculator!B{fy_row}", "=IF(B6<235,\"✓ OK\",\"❌ Overstressed\")"),
        ("Utilization", f"=Calculator!B{util_row}", "< 1.0", "=IF(B7<1,\"✓ OK\",\"❌ Overstressed\")"),
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
    create_fixed_chimney_calculator()