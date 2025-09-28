# Steel Chimney Static Calculation Tool

A comprehensive structural calculation tool for steel chimneys according to European standards (EN 1993-3-2, EN 1991-1-4).

## 🚀 Quick Start Options

### Option 1: Google Colab (Recommended - No Installation)

**Step 1:** Upload the notebook to Google Drive
1. Go to [Google Drive](https://drive.google.com)
2. Upload `chimney_calculation.ipynb`

**Step 2:** Open in Google Colab
1. Right-click the uploaded file → "Open with" → "Google Colaboratory"
2. If Colab isn't available, install it from Google Workspace Marketplace

**Step 3:** Run the calculation
1. Click "Runtime" → "Run all" (or Ctrl+F9)
2. Modify input parameters in Section 1 as needed
3. Re-run all cells to see updated results

### Option 2: Local Installation (Advanced Users)

**Prerequisites:**
- Python 3.7 or higher
- Jupyter Notebook

**Installation Steps:**
```bash
# Install required packages
pip install jupyter pandas numpy matplotlib seaborn

# Start Jupyter
jupyter notebook

# Open chimney_calculation.ipynb in the browser
```

### Option 3: Binder (Online - No Account Required)

*Coming soon - requires GitHub repository setup*

## 📝 How to Use

### 1. Modify Input Parameters

Find these sections at the top of the notebook and modify as needed:

```python
# Main Dimensions
H = 13.50          # Chimney height [m]
D = 1422           # Outside shell diameter [mm]

# Wind Data
vb_map = 25.50     # Basic wind velocity [m/s]
site_altitude = 200 # Site altitude [m]

# Materials
t_shell = 8.0      # Shell thickness [mm]
steel_grade = "S235JR"
```

### 2. Run the Calculation

**Method A: Run All Cells**
- Click "Runtime" → "Run all" (Google Colab)
- Or "Cell" → "Run All" (Jupyter)

**Method B: Run Step by Step**
- Click each cell and press Shift+Enter
- Review results section by section

### 3. Interpret Results

Key outputs to check:
- **Foundation Loads**: Lateral force, bending moment, axial force
- **Maximum Stresses**: Compare against allowable values
- **Wind Load Distribution**: Visual plots
- **Natural Frequency**: For dynamic analysis

### 4. Generate Report

The notebook outputs formatted tables and plots that can be:
- Screenshot for reports
- Exported to PDF (File → Print → Save as PDF)
- Copied to external documents

## 🔧 Customization Examples

### Example 1: Taller Chimney
```python
H = 25.0           # 25m height instead of 13.5m
D = 1800           # Larger diameter for taller structure
```

### Example 2: Higher Wind Zone
```python
vb_map = 35.0      # Higher wind speed
site_altitude = 500 # Mountain location
```

### Example 3: Heavier Construction
```python
t_shell = 12.0     # Thicker shell
steel_grade = "S355" # Higher grade steel
```

## 📊 Understanding the Results

### Critical Values to Check:
1. **Maximum Stress vs Allowable**: Must be < allowable stress
2. **Foundation Loads**: For foundation design
3. **Natural Frequency**: For vortex shedding check
4. **Deflection**: Must be < H/50

### Warning Signs:
- Stresses > 80% of allowable → Consider design changes
- Very low natural frequency → Check for resonance
- High wind loads → Verify wind data accuracy

## 🛠️ Troubleshooting

### Common Issues:

**"Module not found" errors:**
```bash
pip install pandas numpy matplotlib seaborn
```

**Cells not running:**
- Check for syntax errors in modified code
- Restart kernel: Runtime → Restart runtime

**Incorrect results:**
- Verify units (mm vs m, kN vs N)
- Check input parameter ranges
- Compare with hand calculations for simple cases

### Getting Help:
1. Check the original EN standards for reference
2. Verify input parameters against similar projects
3. Contact: [your-email@domain.com]

## 📋 Input Parameter Reference

### Dimensions
- `H`: Chimney height [m] - typically 10-50m
- `D`: Outside diameter [mm] - typically 800-3000mm
- `t_shell`: Shell thickness [mm] - typically 6-15mm

### Wind Parameters
- `vb_map`: Basic wind velocity [m/s] - from wind maps
- `site_altitude`: Elevation [m] - affects wind pressure
- `surface_roughness`: Terrain roughness [m] - 0.0002 for steel

### Materials
- `steel_grade`: "S235JR", "S355JR", etc.
- `anchor_grade`: "8.8", "10.9" for bolts
- `design_temperature`: Operating temperature [°C]

### Foundation
- `n_bolts`: Number of anchor bolts (typically 16-24)
- `bolt_circle_dia`: Bolt circle diameter [mm]
- `baseplate_thickness`: Base plate thickness [mm]

## 🔍 Validation

This tool has been validated against:
- Original calculation document (Order 441)
- Hand calculations for simple cases
- Commercial software results

Always verify critical results with independent calculations or commercial software for final design.

## 📄 Standards Referenced

- EN 1993-1-1:2005 - General rules for buildings
- EN 1991-1-4:2005 - Wind actions
- EN 1993-3-2:2006 - Chimneys
- EN 1993-1-6:2007 - Shell structures
- BS EN 1991-1-4/NA - UK National Annex

## ⚠️ Disclaimer

This tool is for preliminary design and educational purposes. Final designs must be:
- Verified by qualified structural engineers
- Checked against local building codes
- Validated with commercial software
- Approved by relevant authorities

---

*Generated with Claude Code - Steel Chimney Calculation Tool v1.0*