# Volume Accuracy - Mix Generation System

## Project Overview

The **Volume Accuracy** system is a sophisticated data processing application designed for STELLANTIS automotive manufacturing operations. It automates the calculation of vehicle production volume mixes by processing complex multi-value configurations, packet codes, and manufacturing plant specifications.

### Purpose
This system processes automotive production data to:
- Calculate accurate volume mixes for different vehicle configurations
- Handle inclusion/exclusion logic for packet codes
- Process multi-value attributes (engine types, trim levels, transmission, etc.)
- Generate consolidated production planning reports

---

## Table of Contents

1. [System Architecture](#system-architecture)
2. [Features](#features)
3. [Technical Specifications](#technical-specifications)
4. [Installation](#installation)
5. [Usage Guide](#usage-guide)
6. [Data Processing Pipeline](#data-processing-pipeline)
7. [Input File Requirements](#input-file-requirements)
8. [Output Specifications](#output-specifications)
9. [Core Algorithms](#core-algorithms)
10. [Configuration Mappings](#configuration-mappings)
11. [Error Handling](#error-handling)
12. [Performance Optimization](#performance-optimization)
13. [Troubleshooting](#troubleshooting)
14. [Version History](#version-history)

---

## System Architecture

### Component Structure

```
Volume_Accuracy/
├── Volume_Accuracy.py       # Main application with GUI
├── Volume_Parten.py         # Parten-specific processing module
├── optimiz.py               # Optimized version (secondary)
├── History1.py              # Legacy/reference implementation
├── Volume_Accuracy.spec     # PyInstaller specification for executable build
├── assets/
│   └── Vlc_img.png         # STELLANTIS logo for GUI
└── Map Opeperation/
    ├── Griglia.xlsx         # Master reference data (grid)
    ├── Relatorio_61.xlsx    # Input production report
    └── tb_de_para.xlsx      # Multi-value mapping table
```

### Technology Stack

- **Python 3.x**
- **GUI Framework**: Tkinter (native Python GUI)
- **Data Processing**: 
  - `pandas` - DataFrame operations
  - `numpy` - Numerical computations
  - `openpyxl` - Excel file handling
- **Image Processing**: Pillow (PIL)
- **Packaging**: PyInstaller (for standalone .exe)

---

## Features

### Core Capabilities

1. **Multi-Value Processing**
   - Handles complex automotive configuration codes
   - Processes inclusion (+) and exclusion (-) operators
   - Year-based filtering (MY - Model Year)
   - Market-specific configurations

2. **Volume Calculations**
   - Minimum volume logic for included packets
   - Maximum volume logic for excluded packets
   - Fallback to head volume when primary calculations fail
   - Aggregation with deduplication

3. **Plant-Specific Logic**
   - Special handling for FIAPE, FIASA plants
   - Model-specific translations (226, 291, 281, 521, 598, 551)
   - Level mappings (liv.0-13 to LL0-13)

4. **User Interface**
   - Folder selection dialogs
   - Real-time progress tracking
   - Execution time monitoring
   - Detailed logging

---

## Technical Specifications

### System Requirements

**Minimum Requirements:**
- Windows 10/11 (PowerShell compatible)
- 4 GB RAM
- 500 MB free disk space
- Microsoft Excel 2016+ (for viewing outputs)

**Recommended:**
- 8 GB RAM
- SSD storage
- Multi-core processor

### Dependencies

```python
# Core Dependencies
pandas>=1.5.0
numpy>=1.23.0
openpyxl>=3.0.10
Pillow>=9.0.0

# Built-in Libraries
tkinter (included with Python)
threading (included with Python)
re (included with Python)
os (included with Python)
time (included with Python)
datetime (included with Python)
sys (included with Python)
```

---

## Installation

### Option 1: Python Environment

```powershell
# Clone or download the project
cd c:\Users\perna\Desktop\STALLANTIS\Volume_Accuracy

# Install dependencies
pip install pandas numpy openpyxl Pillow

# Run the application
python Volume_Accuracy.py
```

### Option 2: Standalone Executable

```powershell
# Build using PyInstaller
pyinstaller --onefile --noconsole --icon "C:/Users/perna/Desktop/STALLANTIS/Volume_Accuracy/icon.ico"  --add-data  "assets;assets" Volume_Accuracy.py

# Run the executable
.\dist\Volume_Accuracy.exe
```

---

## Usage Guide

### Step-by-Step Process

1. **Launch Application**
   ```powershell
   python Volume_Accuracy.py
   ```

2. **Select Input Folder**
   - Click "Select Folders and Run"
   - Choose folder containing:
     - `Relatorio_61.xlsx`
     - `Griglia.xlsx`
     - `tb_de_para.xlsx`

3. **Select Output Folder**
   - Choose destination for `Volume_Accuracy.xlsx`

4. **Monitor Progress**
   - Progress bar shows completion percentage
   - Log window displays real-time processing steps

5. **Review Results**
   - Output file: `Volume_Accuracy.xlsx`
   - Execution time displayed upon completion

---

## Data Processing Pipeline

### Workflow Diagram

```
┌─────────────────────┐
│  Load Input Files   │
│  - Relatorio_61     │
│  - Griglia          │
│  - tb_de_para       │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│  Data Cleaning      │
│  - Remove .0 suffix │
│  - Filter empty     │
│  - Normalize text   │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│ Process Volumes     │
│  - Included codes   │
│  - Excluded codes   │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│  Merge DataFrames   │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│ Multi-Value Parse   │
│  - Normalize tokens │
│  - Apply mappings   │
│  - Filter MY dates  │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│ Volume Mapping      │
│  - Map included     │
│  - Map excluded     │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│ Final Calculations  │
│  - Volume_Mix       │
│  - Mix percentage   │
│  - Aggregation      │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│  Export Results     │
│  Volume_Accuracy.xlsx│
└─────────────────────┘
```

---

## Input File Requirements

### 1. Relatorio_61.xlsx

**Sheet Name:** `61`

**Required Columns:**
- `Modelo` - Vehicle model code
- `PN` - Part number
- `Plant` - Manufacturing plant code
- `multivalues` - Configuration attributes
- `included` - Included packet codes (comma-separated)
- `excluded` - Excluded packet codes (comma-separated)

**Format:**
```
Modelo | PN    | Plant | multivalues           | included | excluded
226    | 12345 | FIAPE | MT(D,G)+,CC(1.0,1.3)+ | 1,2,3   | 45,46
```

### 2. Griglia.xlsx

**Required Columns:**
- `Model` - Model identifier
- `Plant` - Plant code
- `SINCOM` - Configuration identifier
- `Packet` - Packet name
- `Code` - Packet code
- `included` - Included configurations
- `Multivalues` - Associated multi-values
- `Volume Head` - Header volume
- `Volume TT` - Total volume
- `Volume` - Individual packet volume

### 3. tb_de_para.xlsx

**Sheet Name:** `Coded`

**Required Columns:**
- `MultiValues` - Raw token (e.g., "MT(D)")
- `Resp.1` - Primary translation
- `Resp.2` - Secondary translation
- `Griglia Italiano` - Italian packet name
- `Griglia Inglês` - English packet name

---

## Output Specifications

### Volume_Accuracy.xlsx

**Output Columns:**

| Column | Type | Description |
|--------|------|-------------|
| `Modelo` | String | Vehicle model code |
| `PN` | String | Part number |
| `Plant` | String | Manufacturing plant |
| `multivalues` | String | Configuration attributes |
| `included` | String | Included packet codes |
| `excluded` | String | Excluded packet codes |
| `Unique_Key` | String | Row identifier |
| `VolumeTT` | Float | Total volume |
| `Volume_Mix` | Float | Calculated mix volume |
| `Mix` | Float | Mix percentage (Volume_Mix/VolumeTT) |

**Calculation Logic:**
```python
# Standard calculation
Volume_Mix = max(Final_Volume_include - Final_Volume_Excluded, 0)

# Fallback when included fails
Volume_Mix = volume_Head - Code_excluded

# When all fields empty
Volume_Mix = VolumeTT
```

---

## Core Algorithms

### 1. Multi-Value Normalization

**Purpose:** Extract and filter configuration tokens from raw multi-value strings.

**Algorithm:**
```python
def normalize_multivalues(raw):
    """
    Input:  "MT(D,G)+,CC(1.0,1.3)+,MY(26,27)-"
    Output: ["MT(D)+", "MT(G)+", "CC(1.0)+", "CC(1.3)+"]
    
    Notes:
    - Filters out current/next model years
    - Splits comma-separated values
    - Preserves +/- operators
    """
```

**Steps:**
1. Extract tokens using regex: `(\w+)\(([^)]+)\)([+-])`
2. Filter Model Year (MY) tokens for current/next year
3. Split multi-value lists within parentheses
4. Return normalized list

### 2. Packet Volume Computation

**Purpose:** Calculate min/max volumes based on packet codes.

**Modes:**
- **MIN Mode** (included packets): All packets must match; returns 0 if any missing
- **MAX Mode** (excluded packets): Returns maximum of available matches

**Algorithm:**
```python
def compute_volume_metric(packet_values, sincom, filtered_griglia, mode):
    """
    1. Normalize packet codes (pad with zeros: "1" → "001")
    2. Match by Code column (exact/contains)
    3. Match by Packet_cleaned column (packet names)
    4. Clean and convert volumes to float
    5. Return min/max based on mode
    """
```

### 3. Volume Calculation Priority

**Hierarchy:**
```
1. Code_included (direct packet match)
2. multi_included_min_volume (multi-value mapping)
3. Final_Volume_include = 
   - If both non-zero AND multivalues present: min(Code, Multi)
   - If included fails (0): 0
   - Otherwise: max(Code, Multi)

4. Fallback:
   If Final_Volume_include = 0 AND included empty AND excluded present:
   Final_Volume_include = volume_Head - Code_excluded
```

---

## Configuration Mappings

### Level Mappings (FIAPE Plant)

**Models 226, 291, 281, 521:**
```python
livello_map = {
    "liv.0": "LL0",   "liv.1": "LL1",   "liv.2": "LL2",
    "liv.3": "LL3",   "liv.4": "LL4",   "liv.5": "LL5",
    "liv.6": "LL6",   "liv.7": "LL7",   "liv.8": "LL8",
    "liv.9": "LL9",   "liv.10": "LL10", "liv.11": "LL11",
    "liv.12": "LL12", "liv.13": "LL13"
}
```

**Models 598, 551:**
```python
map_598 = {
    "liv.0": "Level 0",  "liv.1": "Level 1",  # ...
    "LL0": "Level 0",    "LL1": "Level 1",    # ...
}
```

### Market Codes

```python
markets = [
    "ARGENTINA", "BRASILE", "MESSICO", "ALTRI MERCATI",
    "BRAZIL", "MEXICO", "OTHER MARKETS"
]
```

**Special Handling:**
- Market tokens use `Volume` column instead of `Volume Head`
- Market matches use max() instead of min()

---

## Error Handling

### Validation Checks

1. **Missing Files**
   ```python
   if not os.path.exists(relatorio_file):
       raise FileNotFoundError("Missing required files")
   ```

2. **Empty Modelo Column**
   ```python
   mask = ~df_61['Modelo'].str.strip().str.lower().isin(['', 'none', 'nan'])
   df_61 = df_61[mask]
   ```

3. **Invalid Sheet Names**
   ```python
   if "61" not in wb_61.sheetnames:
       raise ValueError("Sheet '61' not found")
   ```

4. **Numeric Conversion**
   ```python
   df_61["Volume_Mix"] = pd.to_numeric(df_61["Volume_Mix"], errors="coerce")
   ```

### Logging System

**Thread-Safe GUI Logging:**
```python
def log_message(message):
    print(message)  # Console
    root.after(0, lambda: log_widget.insert(tk.END, message + "\n"))  # GUI
```

**Progress Tracking:**
```python
update_progress(10, "Processing included codes...")
update_progress(60, "Mapping excluded multivalues...")
```

---

## Performance Optimization

### Optimization Techniques

1. **Read-Only Excel Loading**
   ```python
   wb_61 = load_workbook(filename=File_Rela_61, read_only=True, data_only=True)
   ```

2. **Early Exit for Empty Data**
   ```python
   if not raw or raw.lower() == 'none':
       return ("", "", "")
   ```

3. **Pre-compiled Regex**
   ```python
   pattern = re.compile(r"(\w+)\(([^)]+)\)([+-])")
   ```

4. **Vectorized Operations**
   ```python
   df_61["Final_Volume_Excluded"] = np.maximum(
       df_61["Code_excluded"], 
       df_61["multi_excluded_max_volume"]
   )
   ```

5. **Set Operations for Deduplication**
   ```python
   unique_packet_values = set(packet_values)
   ```

### Typical Performance

| Data Size | Processing Time |
|-----------|-----------------|
| 1,000 rows | ~30 seconds |
| 5,000 rows | ~2 minutes |
| 10,000 rows | ~5 minutes |

---

## Troubleshooting

### Common Issues

#### 1. Missing Dependencies
**Error:** `ModuleNotFoundError: No module named 'pandas'`

**Solution:**
```powershell
pip install pandas numpy openpyxl Pillow
```

#### 2. Image File Not Found
**Error:** `FileNotFoundError: assets/Vlc_img.png`

**Solution:**
- Ensure `assets/` folder exists
- Verify `Vlc_img.png` is present
- For PyInstaller builds, check `--add-data` flag

#### 3. Empty Output
**Symptom:** `Volume_Mix` column all zeros

**Diagnosis:**
- Check if `Modelo` column has valid data
- Verify `Griglia.xlsx` has matching Model/Plant combinations
- Review `included`/`excluded` packet codes

#### 4. Incorrect Mix Values
**Symptom:** Mix percentages > 1.0 or negative

**Diagnosis:**
- Verify `VolumeTT` is not zero
- Check for data type issues (strings vs numbers)
- Review fallback logic application

---

## Version History

### Version 2.1 (Current - Volume_Parten.py & Executable Build)
**Date:** January 2026

**New Features:**
- ✅ Added Volume_Parten.py module for dedicated parten data processing
- ✅ Included PyInstaller spec file (Volume_Accuracy.spec) for streamlined executable builds
- ✅ Enhanced build process with proper asset inclusion

**Improvements:**
- Minor code cleanup and optimization in processing modules

### Version 2.0 (optimiz.py & Volume_Accuracy.py)
**Date:** November 2025

**Major Changes:**
- ✅ Fixed `.0` suffix removal for numeric columns (Modelo, PN)
- ✅ Improved packet code normalization (pad with zeros)
- ✅ Enhanced market handling logic
- ✅ Added fallback to `volume_Head - Code_excluded`
- ✅ Optimized multi-value processing (early exit for empty values)
- ✅ Better SINCOM matching for multi-included volumes
- ✅ Refined minimum/maximum volume calculation logic

**Bug Fixes:**
- Fixed duplicate SINCOM entries
- Corrected level mapping for FIASA plant
- Resolved packet reference tracking

### Version 1.0 (History1.py)
**Date:** Legacy

**Features:**
- Basic volume processing
- GUI with folder selection
- Multi-value parsing
- Excel output generation

---

## Data Flow Examples

### Example 1: Standard Included/Excluded Processing

**Input Row:**
```
Modelo: 226
Plant: FIAPE
included: 1,2,3
excluded: 45,46
multivalues: MT(D)+,CC(1.0)+
SINCOM: ABC123
```

**Processing:**
1. Find packets 1,2,3 in Griglia → Code_included = min(vol1, vol2, vol3)
2. Find packets 45,46 in Griglia → Code_excluded = max(vol45, vol46)
3. Parse multivalues → Multi_included: "mt(d),cc(1.0)"
4. Map to Griglia volumes → multi_included_min_volume = min(volMT, volCC)
5. Calculate:
   ```python
   Final_Volume_include = min(Code_included, multi_included_min_volume)
   Final_Volume_Excluded = max(Code_excluded, multi_excluded_max_volume)
   Volume_Mix = max(Final_Volume_include - Final_Volume_Excluded, 0)
   ```

### Example 2: Fallback Logic

**Input Row:**
```
Modelo: 291
Plant: FIAPE
included: (empty)
excluded: 10
multivalues: (empty)
volume_Head: 1000
```

**Processing:**
1. Code_included = 0 (no packets)
2. multi_included_min_volume = NaN (no multivalues)
3. Final_Volume_include = 0
4. **Trigger Fallback:**
   ```python
   Final_Volume_include = volume_Head - Code_excluded
   Final_Volume_include = 1000 - 50 = 950
   Used_Fallback_HeadVolume = True
   ```

---

## File Format Specifications

### Excel Cell Formatting

**Relatorio_61.xlsx:**
- Text columns: General format
- Numeric columns (Modelo, PN): Remove `.0` suffix during load
- Multi-value column: Text with special characters

**Griglia.xlsx:**
- Volume columns: Numeric (handle commas as thousands separators)
- Code column: Text (preserve leading zeros)

### Special Characters Handling

**Supported in multivalues:**
- Parentheses: `()`
- Plus/Minus: `+-`
- Commas: `,` (value separator)
- Periods: `.` (decimal separator)

**Reserved Prefixes:**
- `MT` - Transmission type
- `CC` - Engine displacement
- `MY` - Model Year
- `EC` - Ecology level
- `TR` - Trim level

---

## Advanced Features

### 1. Thread-Safe GUI Updates

**Challenge:** Updating GUI from background threads causes crashes.

**Solution:**
```python
root.after(0, lambda: (
    log_widget.config(state=tk.NORMAL),
    log_widget.insert(tk.END, message + "\n"),
    log_widget.config(state=tk.DISABLED)
))
```

### 2. Dynamic Column Preservation

**Ensures all original columns are retained:**
```python
preserved_cols = list(df_61.columns) + ["SINCOM", "volume_Head", "VolumeTT", "Code_excluded"]
for col in preserved_cols:
    if col not in result_df.columns:
        result_df[col] = np.nan
```

### 3. Packet Code Normalization

**Handles variable-length codes:**
```python
if len(packet_code) == 1:
    packet_code = "00" + packet_code  # "1" → "001"
elif len(packet_code) == 2:
    packet_code = "0" + packet_code   # "45" → "045"
```

---

## Configuration Files

### PyInstaller Spec File (for building .exe)

```python
# Volume_Accuracy.spec
a = Analysis(
    ['Volume_Accuracy.py'],
    datas=[('assets', 'assets')],
    hiddenimports=['pandas', 'numpy', 'openpyxl', 'PIL'],
    hookspath=[],
    runtime_hooks=[],
)
pyz = PYZ(a.pure, a.zipped_data)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    name='Volume_Accuracy',
    icon='assets/icon.ico',
    console=False,
)
```

---

## API Reference (Key Functions)

### Data Loading
```python
def load_dataframes(File_Rela_61, File_Griglia)
    """
    Loads Excel files into pandas DataFrames with cleaning.
    Returns: (df_61, df_griglia)
    """
```

### Volume Processing
```python
def process_volume_table(df_61, df_griglia, field, min_col_name)
    """
    Processes included/excluded packet volumes.
    field: 'included' or 'excluded'
    Returns: DataFrame with volume columns
    """
```

### Multi-Value Handling
```python
def normalize_multivalues(raw)
    """
    Normalizes multi-value tokens, filters MY dates.
    Returns: List of normalized tokens
    """

def extract_and_save_structured_data(df_61, mapping_file_path, df_griglia)
    """
    Main orchestrator for multi-value processing and volume calculation.
    Saves final output Excel file.
    """
```

### Mapping Functions
```python
def map_multi_included_to_griglia(df_flattened, df_griglia)
    """
    Maps multi-value included tokens to Griglia volumes (MIN logic).
    Returns: DataFrame with 'multi_included_min_volume' column
    """

def map_multi_excluded_to_griglia(df_flattened, griglia_path, df_mapping)
    """
    Maps multi-value excluded tokens to Griglia volumes (MAX logic).
    Returns: DataFrame with 'multi_excluded_max_volume' column
    """
```

---

## Best Practices

### For Users
1. ✅ Always back up input files before processing
2. ✅ Verify Model/Plant codes match between Relatorio and Griglia
3. ✅ Use consistent date formats in MY() tokens
4. ✅ Review log output for warnings

### For Developers
1. ✅ Use `.copy()` when modifying DataFrames to avoid SettingWithCopyWarning
2. ✅ Convert columns to numeric with `errors='coerce'` for safety
3. ✅ Apply string operations with `.astype(str)` first
4. ✅ Use `root.after()` for all GUI updates from threads
5. ✅ Add progress updates every 5-15% of workflow

---

## Support & Contact

**Project Owner:** STELLANTIS Manufacturing Operations  
**Repository:** c:\Users\perna\Desktop\STALLANTIS\Volume_Accuracy  
**Branch:** version-two  

**For Issues:**
1. Check [Troubleshooting](#troubleshooting) section
2. Review log output in GUI
3. Verify input file formats
4. Check intermediate Excel files (Befoer_multi.xlsx, AFter_Included.xlsx)

---

## License

**Proprietary - STELLANTIS Internal Use Only**

This software is developed for internal STELLANTIS manufacturing operations. Unauthorized distribution or use outside STELLANTIS organization is prohibited.

---

## Glossary

| Term | Definition |
|------|------------|
| **SINCOM** | Configuration identifier for vehicle specifications |
| **Griglia** | Master reference grid/table (Italian: "grid") |
| **Packet Code** | Numeric code identifying configuration packages |
| **Multi-Value** | Complex attribute with multiple parameters (e.g., MT(D,G)+) |
| **Mix** | Ratio of Volume_Mix to total volume (VolumeTT) |
| **Volume Head** | Header/reference volume for a configuration |
| **Volume TT** | Total volume for a specific model/plant/SINCOM combination |
| **LL** | Level designation (LL0-LL13) |
| **MY** | Model Year |
| **FIAPE** | Plant code (example manufacturing plant) |
| **FIASA** | Plant code (example manufacturing plant) |

---

## Appendix

### A. Sample Data Structures

**Relatorio_61.xlsx Row:**
```
{
  "Modelo": "226",
  "PN": "12345",
  "Plant": "FIAPE",
  "multivalues": "MT(D)+,CC(1.0)+,MY(25)-",
  "included": "1,2,3",
  "excluded": "45,46",
  "Unique_Key": "226_FIAPE_12345"
}
```

**Griglia.xlsx Row:**
```
{
  "Model": "226",
  "Plant": "FIAPE",
  "SINCOM": "ABC123",
  "Packet": "Pacchetto Base",
  "Code": "001",
  "included": "1,2,3",
  "Multivalues": "MT(D),CC(1.0)",
  "Volume Head": 1000.0,
  "Volume TT": 5000.0,
  "Volume": 150.0
}
```

### B. Regex Patterns

**Multi-Value Token Extraction:**
```regex
(\w+)\(([^)]+)\)([+-])
```
- Group 1: Prefix (MT, CC, MY, etc.)
- Group 2: Values (D,G or 1.0,1.3)
- Group 3: Operator (+ or -)

**Level Detection:**
```regex
liv\.?\s*\d
```
Matches: `liv.5`, `liv5`, `liv. 10`

---

## End of Documentation

**Last Updated:** January 13, 2026  
**Document Version:** 2.1  
**Author:** Technical Documentation Team

---

*This documentation is subject to updates as the system evolves. Please check repository for the latest version.*
