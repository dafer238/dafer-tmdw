"""
OXT Package Creator for CoolProp LibreOffice Add-in
Creates a .oxt extension package that can be installed in LibreOffice Calc
"""

from pathlib import Path
import zipfile
import os

# Get the directory where this script is located
script_dir = Path(__file__).parent

# Paths
output_oxt_path = script_dir / "CoolPropLibre.oxt"
python_script_path = script_dir / "CoolPropWrapper_Simple.py"
basic_script_path = script_dir / "CoolProp.xba"

# Verify files exist
if not python_script_path.exists():
    print(f"Error: Python script not found at {python_script_path}")
    exit(1)

if not basic_script_path.exists():
    print(f"Error: Basic script not found at {basic_script_path}")
    exit(1)

# Create META-INF/manifest.xml
manifest_xml = """<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
    <manifest:file-entry manifest:media-type="application/vnd.sun.star.basic-library" 
                         manifest:full-path="CoolProp/"/>
</manifest:manifest>
"""

# Create Basic library structure
library_xml = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">
<library:library xmlns:library="http://openoffice.org/2000/library" library:name="CoolProp" library:readonly="false" library:passwordprotected="false">
 <library:element library:name="CoolProp"/>
</library:library>
"""

dialog_xml = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">
<library:library xmlns:library="http://openoffice.org/2000/library" library:name="CoolProp" library:readonly="false" library:passwordprotected="false"/>
"""

script_xml = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="CoolProp" script:language="StarBasic" script:moduleType="normal"/>
"""

# Create description.xml
description_xml = """<?xml version="1.0" encoding="UTF-8"?>
<description xmlns="http://openoffice.org/extensions/description/2006" 
             xmlns:xlink="http://www.w3.org/1999/xlink">
    <identifier value="org.coolprop.calc.addin"/>
    <version value="2.0.0"/>
    <display-name>
        <name lang="en">CoolProp Thermodynamic Properties</name>
    </display-name>
    <extension-description>
        <src lang="en" xlink:href="description-en.txt"/>
    </extension-description>
    <dependencies>
        <OpenOffice.org-minimal-version value="4.0" dep:name="OpenOffice.org 4.0"/>
    </dependencies>
</description>
"""

# Create description text
description_text = """CoolProp Wrapper for LibreOffice Calc

This extension provides thermodynamic property calculations using the CoolProp library.

Functions available:
- CPROP, CPROP_E, CPROP_SI - Real fluid properties
- CPROPHA, CPROPHA_E, CPROPHA_SI - Humid air properties

Supports both SI units (K, Pa, J/kg) and engineering units (°C, bar, kJ/kg).

Requirements:
- CoolProp Python package must be installed in LibreOffice's Python environment

For more information, see README_LibreOffice.md
"""

# Create the .oxt package
print(f"Creating OXT package at: {output_oxt_path}")

with zipfile.ZipFile(output_oxt_path, 'w', zipfile.ZIP_DEFLATED) as oxt:
    # Add manifest
    oxt.writestr("META-INF/manifest.xml", manifest_xml)
    
    # Add description files
    oxt.writestr("description.xml", description_xml)
    oxt.writestr("description-en.txt", description_text)
    
    # Add the Python script in Scripts folder (so Basic can find it)
    oxt.write(python_script_path, arcname="Scripts/python/CoolPropWrapper_Simple.py")
    
    # Add Basic library files
    oxt.writestr("CoolProp/script.xlb", library_xml)
    oxt.writestr("CoolProp/dialog.xlb", dialog_xml)
    oxt.writestr("CoolProp/CoolProp.xml", script_xml)
    oxt.write(basic_script_path, arcname="CoolProp/CoolProp.xba")
    
    print(f"  Added: Scripts/python/CoolPropWrapper_Simple.py")
    print(f"  Added: CoolProp Basic library")
    print(f"  Added: META-INF/manifest.xml")
    print(f"  Added: description.xml")
    print(f"  Added: description-en.txt")

print(f"\nSuccessfully created: {output_oxt_path.name}")
print("\nTo install:")
print("1. Open LibreOffice Calc")
print("2. Go to Tools → Extension Manager")
print("3. Click 'Add' and select the .oxt file")
print("4. Restart LibreOffice")
