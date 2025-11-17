# CoolProp Wrapper for LibreOffice Calc

This directory contains the LibreOffice Calc add-in implementation of CoolProp thermodynamic property calculations.

## Quick Start (LibreOffice Portable)

If you're using **LibreOffice Portable** and getting the `No module named 'CoolProp'` error:

1. **Install CoolProp** using the automated script:
   ```powershell
   cd Libreoffice
   .\install_coolprop_portable.ps1
   ```
   
   Or manually follow the [Method B installation steps](#method-b-manual-installation-libreoffice-portable-or-when-pip-is-not-available) below.

2. **Install the extension**:
   - **IMPORTANT**: If you previously installed CoolProp extension, remove it first:
     - Go to **Tools → Extension Manager**
     - Select any old CoolProp extension and click **Remove**
     - Close LibreOffice completely
   - Open LibreOffice Calc
   - Go to **Tools → Extension Manager**
   - Click **Add** and select the new `CoolPropLibre.oxt`
   - Accept all prompts
   - **Close LibreOffice completely** (all windows, check system tray/taskbar)
   - Reopen LibreOffice Calc

3. **Verify installation**:
   - Go to **Tools → Macros → Organize Macros → Basic**
   - You should see "CoolProp" library in the list
   - Close the dialog

4. **Use the functions**:
   ```
   =CPROP("H", "T", 25, "P", 1.01325, "Water")
   ```
   Should return approximately **104.92** kJ/kg

## Installation

### Prerequisites

1. **LibreOffice Calc** (version 6.0 or higher recommended)
2. **Python 3** (LibreOffice includes its own Python interpreter)
3. **CoolProp Python Package**

### Installing CoolProp for LibreOffice's Python

LibreOffice comes with its own Python interpreter. You need to install CoolProp into LibreOffice's Python environment.

**IMPORTANT**: LibreOffice Portable and some installations don't include pip, so you'll need to install CoolProp manually.

#### Method A: Using pip (if available)

##### On Windows (Standard Installation):
```bash
# Find LibreOffice's Python (usually in C:\Program Files\LibreOffice\program\)
cd "C:\Program Files\LibreOffice\program"
python.exe -m pip install CoolProp
```

##### On Linux:
```bash
# LibreOffice Python is usually at /usr/lib/libreoffice/program/python
/usr/lib/libreoffice/program/python -m pip install CoolProp
```

##### On macOS:
```bash
# LibreOffice Python path on macOS
/Applications/LibreOffice.app/Contents/Resources/python -m pip install CoolProp
```

#### Method B: Manual Installation (LibreOffice Portable or when pip is not available)

1. **Download CoolProp wheel** for your LibreOffice's Python version:
   ```bash
   # First, check your LibreOffice Python version (32-bit or 64-bit, Python version)
   # On Windows, LibreOffice Portable typically uses Python 3.10 32-bit
   
   # Download the appropriate wheel (using your system Python with pip)
   pip download CoolProp --dest "C:\Temp" --only-binary :all: --platform win32 --python-version 310 --no-deps
   ```

2. **Extract the wheel to LibreOffice's site-packages**:
   
   For **LibreOffice Portable** on Windows:
   ```bash
   # Extract the wheel (it's just a zip file)
   # Replace paths with your actual LibreOffice Portable location
   python -m zipfile -e "C:\Temp\coolprop-7.2.0-cp310-cp310-win32.whl" "C:\...\LibreOfficePortable\App\libreoffice\program\python-core-3.10.18\lib\site-packages"
   ```
   
   For **Standard LibreOffice** on Windows:
   ```bash
   python -m zipfile -e "C:\Temp\coolprop-7.2.0-cp310-cp310-win_amd64.whl" "C:\Program Files\LibreOffice\program\python-core-3.10.18\lib\site-packages"
   ```

3. **Verify installation**:
   ```bash
   # Run LibreOffice's Python
   "C:\...\LibreOffice\program\python.exe" -c "import CoolProp; print(CoolProp.__version__)"
   ```
   
   Should output: `7.2.0` (or similar version number)

#### Quick Reference for Common LibreOffice Portable Paths:

- **Python executable**: `...\LibreOfficePortable\App\libreoffice\program\python.exe`
- **Site-packages**: `...\LibreOfficePortable\App\libreoffice\program\python-core-3.10.18\lib\site-packages`
- **Check Python version**: Run `python.exe -c "import sys; print(sys.version)"`

### Installing the Add-in

#### Method 1: Using the .oxt Extension (Recommended)

1. Open LibreOffice Calc
2. Go to **Tools → Extension Manager**
3. Click **Add** button
4. Navigate to and select `CoolPropLibre.oxt`
5. Click **Accept** to accept the license
6. Restart LibreOffice Calc

#### Method 2: Manual Python Script Installation

1. Copy `CoolPropWrapper_Libreoffice.py` to LibreOffice's user script directory:
   - **Windows**: `%APPDATA%\LibreOffice\4\user\Scripts\python\`
   - **Linux**: `~/.config/libreoffice/4/user/Scripts/python/`
   - **macOS**: `~/Library/Application Support/LibreOffice/4/user/Scripts/python/`

2. Restart LibreOffice Calc

3. The functions should now be available in Calc

## Usage in LibreOffice Calc

### Available Functions

All functions are available with the prefix shown in the function wizard under the **CoolProp** category:

#### Real Fluid Properties

- **`CPROP(output, name1, value1, name2, value2, fluid)`** - Engineering units (default)
- **`CPROP_E(output, name1, value1, name2, value2, fluid)`** - Engineering units (explicit)
- **`CPROP_SI(output, name1, value1, name2, value2, fluid)`** - SI units

#### Humid Air Properties

- **`CPROPHA(output, name1, value1, name2, value2, name3, value3)`** - Engineering units (default)
- **`CPROPHA_E(output, name1, value1, name2, value2, name3, value3)`** - Engineering units (explicit)
- **`CPROPHA_SI(output, name1, value1, name2, value2, name3, value3)`** - SI units

### Examples

#### Calculate water enthalpy (engineering units):
```
=CPROP("H", "T", 25, "P", 1.01325, "Water")
```
Returns enthalpy in kJ/kg at 25°C and 1.01325 bar.

#### Calculate water enthalpy (SI units):
```
=CPROP_SI("H", "T", 298.15, "P", 101325, "Water")
```
Returns enthalpy in J/kg at 298.15 K and 101325 Pa.

#### Calculate humidity ratio:
```
=CPROPHA("W", "T", 25, "P", 1.01325, "RH", 0.5)
```
Returns humidity ratio at 25°C, 1.01325 bar, and 50% RH.

### Unit Systems

#### Engineering Units (CPROP, CPROP_E, CPROPHA, CPROPHA_E):
- Temperature: **°C**
- Pressure: **bar**
- Enthalpy, Internal Energy, Entropy: **kJ/kg**
- Specific Heat: **kJ/kg/K**

#### SI Units (CPROP_SI, CPROPHA_SI):
- Temperature: **K**
- Pressure: **Pa**
- Enthalpy, Internal Energy, Entropy: **J/kg**
- Specific Heat: **J/kg/K**

### Property Aliases

The add-in supports many property name aliases for convenience:

| Standard | Aliases               |
| -------- | --------------------- |
| T        | temp, temperature     |
| P        | pres, pressure        |
| H        | enth, enthalpy, hmass |
| S        | entr, entropy, smass  |
| D        | rho, dens, dmass      |
| Q        | quality, x            |
| MU       | viscosity             |
| K        | conductivity          |

See the main README.md for a complete list of available properties.

## Troubleshooting

### "No module named 'CoolProp'" Error

This is the most common error. It means CoolProp is not installed in LibreOffice's Python environment.

**Solution:**

1. **Find LibreOffice's Python path**:
   - LibreOffice Portable: `...\LibreOfficePortable\App\libreoffice\program\python.exe`
   - Standard Windows: `C:\Program Files\LibreOffice\program\python.exe`
   - Linux: `/usr/lib/libreoffice/program/python`
   - macOS: `/Applications/LibreOffice.app/Contents/Resources/python`

2. **Check Python version and architecture** (32-bit vs 64-bit):
   ```bash
   "path\to\libreoffice\python.exe" -c "import sys; print(sys.version)"
   ```

3. **Install CoolProp using Method B** (Manual Installation) from the installation section above

4. **Verify installation**:
   ```bash
   "path\to\libreoffice\python.exe" -c "import CoolProp; print(CoolProp.__version__)"
   ```
   Should print version number (e.g., `7.2.0`)

5. **Restart LibreOffice completely** and try installing the extension again

### Functions Not Appearing

1. Verify CoolProp is installed in LibreOffice's Python (see above)

2. Check that the extension is installed in **Tools → Extension Manager**

3. Restart LibreOffice completely (close all windows, including quickstarter)

4. Check LibreOffice's user directory for errors

### Import Errors in IDE

If you see "Import CoolProp could not be resolved" errors when editing the Python file in your IDE, these are expected - the imports only work when running in LibreOffice's Python environment.

### Function Returns #NAME? Error

**This means LibreOffice doesn't recognize the function names.** Common causes:

1. **Extension not installed or not enabled**:
   - Go to **Tools → Extension Manager**
   - Check if "CoolProp Thermodynamic Properties" appears in the list
   - If it shows but is disabled, enable it
   - If it's not in the list, install `CoolPropLibre.oxt`

2. **Need to restart LibreOffice**:
   - Close **ALL** LibreOffice windows (including quickstarter)
   - Reopen LibreOffice Calc
   - Try the function again

3. **Remove old extension first**:
   - Go to **Tools → Extension Manager**
   - If you see an old version of CoolProp, select it and click **Remove**
   - Restart LibreOffice
   - Install the new `CoolPropLibre.oxt`
   - Restart LibreOffice again

4. **Function name must be ALL CAPS**:
   - Use `=CPROP(...)` **NOT** `=cprop(...)` or `=CProp(...)`
   - Use `=CPROPHA(...)` **NOT** `=cpropha(...)`

5. **Verify CoolProp is installed** (see "No module named 'CoolProp'" section above)

**After fixing, try this test formula**:
```
=CPROP("H", "T", 25, "P", 1.01325, "Water")
```
Should return approximately 104.92 kJ/kg.

**If you still get errors**, the Python script might not be accessible. Try:
- Go to **Tools → Macros → Edit Macros**
- In the macro editor, go to **Tools → Macros → Organize Macros → Python**
- Check if CoolPropWrapper_Simple.py appears in the list
- If not, the extension may not have installed correctly - try reinstalling

### Function Returns #VALUE! Error

- Check that all numeric arguments are actually numbers, not text
- Verify property names are valid (e.g., "T", "P", "H", "S")
- Ensure fluid name is valid and case-sensitive (e.g., "Water", "R134a", "Air")

### Calculation Errors

- Check that property names are valid CoolProp properties
- Verify units are in the correct system (SI vs Engineering)
- Ensure the fluid name is valid (case-sensitive)
- Check that the state point is valid (not in two-phase region for single-phase properties)
- Try using SI units with `CPROP_SI` or `CPROPHA_SI` to verify the issue isn't unit conversion

### LibreOffice Portable Specific Issues

- LibreOffice Portable does **not** include pip by default
- Use **Method B** (Manual Installation) to install CoolProp
- Make sure you download the correct wheel file (32-bit for most Portable versions)
- Python version is typically 3.10 for recent LibreOffice Portable versions

## Creating the .oxt Extension Package

To rebuild the .oxt file:

1. Update `OXT_creator.py` with correct paths
2. Run: `python OXT_creator.py`
3. The .oxt file will be created with proper manifest and metadata

## Development Notes

The LibreOffice add-in implementation:
- Uses the UNO (Universal Network Objects) interface for LibreOffice integration
- Implements XAddIn, XServiceName, and XServiceInfo interfaces
- Provides function descriptions for the Function Wizard
- Maintains feature parity with the Excel add-in (C# version)
- Supports all the same property aliases and unit conversions

## License

See LICENSE file in the root directory.

## Support

For issues specific to LibreOffice integration, check:
- LibreOffice Python scripting documentation
- CoolProp documentation at http://www.coolprop.org/
