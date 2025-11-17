"""
CoolProp Wrapper for LibreOffice Calc
Provides thermodynamic property calculations using CoolProp library
with support for both SI and Engineering units.
"""

try:
    from CoolProp.CoolProp import PropsSI, HAPropsSI, get_global_param_string
except ImportError:
    # Provide helpful error message if CoolProp is not installed
    def PropsSI(*args):
        return "ERROR: CoolProp not installed. See README_LibreOffice.md"
    def HAPropsSI(*args):
        return "ERROR: CoolProp not installed. See README_LibreOffice.md"
    def get_global_param_string(*args):
        return "CoolProp not installed"

import uno
import unohelper
from com.sun.star.sheet import XAddIn
from com.sun.star.lang import XServiceName, XServiceInfo

# Service implementation name
IMPLE_NAME = "org.openoffice.sheet.addin.CoolPropWrapper"
SERVICE_NAME = "com.sun.star.sheet.AddIn"


def ConvertToSI(name, value):
    conversions = {
        "T": lambda x: x + 273.15,
        "P": lambda x: x * 1e5,
        "H": lambda x: x * 1000,
        "U": lambda x: x * 1000,
        "S": lambda x: x * 1000,
        "Cp": lambda x: x * 1000,
        "Cpmass": lambda x: x * 1000,
        "Cvmass": lambda x: x * 1000,
    }
    return conversions.get(name, lambda x: x)(value)

def ConvertFromSI(name, value):
    conversions = {
        "T": lambda x: x - 273.15,
        "P": lambda x: x / 1e5,
        "H": lambda x: x / 1000,
        "U": lambda x: x / 1000,
        "S": lambda x: x / 1000,
        "Cp": lambda x: x / 1000,
        "Cpmass": lambda x: x / 1000,
        "Cvmass": lambda x: x / 1000,
    }
    return conversions.get(name, lambda x: x)(value)

def FormatName(name):
    """Normalize property name aliases to CoolProp standard names"""
    m = {
        # Temperature
        "t": "T", "temp": "T", "temperature": "T",
        # Pressure
        "p": "P", "pres": "P", "pressure": "P",
        # Enthalpy
        "h": "H", "enth": "H", "enthalpy": "H", "hmass": "H",
        # Internal Energy
        "u": "U", "internalenergy": "U", "umass": "U",
        # Entropy
        "s": "S", "entr": "S", "entropy": "S", "smass": "S",
        # Molar Density
        "dmolar": "Dmolar", "dmol": "Dmolar",
        # Delta
        "delta": "Delta",
        # Density
        "rho": "D", "dens": "D", "dmass": "D",
        # Specific Heat
        "cvmass": "Cvmass", "cv": "Cvmass", 
        "cpmass": "Cpmass", "cp": "Cpmass",
        "cpmolar": "Cpmolar", "cpmol": "Cpmolar",
        "cvmolar": "Cvmolar", "cvmol": "Cvmolar",
        # Quality
        "q": "Q", "quality": "Q", "x": "Q",
        # Additional properties
        "tau": "Tau",
        "alpha0": "Alpha0",
        "alphar": "Alphar",
        "speed_of_sound": "A", "a": "A",
        "bvirial": "Bvirial",
        "k": "K", "conductivity": "K",
        "cvirial": "Cvirial",
        "dipole_moment": "DIPOLE_MOMENT",
        "fh": "FH",
        "g": "G", "gmass": "G",
        "helmoltzmass": "HELMHOLTZMASS",
        "helmholtzmolar": "HELMHOLTZMOLAR",
        "gamma": "gamma",
        "isobaric_expansion_coefficient": "isobaric_expansion_coefficient",
        "isothermal_compressibility": "isothermal_compressibility",
        "surface_tension": "surface_tension",
        "mm": "MM", "molar_mass": "MM",
        "pcrit": "Pcrit", "p_critical": "Pcrit",
        "phase": "Phase",
        "pmax": "pmax",
        "pmin": "pmin",
        "prandtl": "Prandtl",
        "ptriple": "ptriple",
        "p_reducing": "p_reducing",
        "rhocrit": "rhocrit",
        "rhomass_reducing": "rhomass_reducing",
        "smolar_residual": "Smolar_residual",
        "tcrit": "Tcrit",
        "tmax": "Tmax",
        "tmin": "Tmin",
        "ttriple": "Ttriple",
        "t_freeze": "T_freeze",
        "t_reducing": "T_reducing",
        "mu": "MU", "viscosity": "MU",
        "z": "Z"
    }
    return m.get(name.lower(), name)

def FormatName_HA(name):
    """Normalize humid air property name aliases to CoolProp standard names"""
    m = {
        # Wet Bulb Temperature
        "twb": "Twb", "wetbulb": "Twb", "t_wb": "Twb",
        # Dew Point Temperature
        "tdp": "Tdp", "dewpoint": "Tdp", "t_dp": "Tdp",
        # Dry Bulb Temperature
        "t": "T", "tdb": "T", "t_db": "T",
        # Pressure
        "p": "P", 
        "p_w": "P_w",
        # Relative Humidity
        "r": "R", "rh": "R", "relhum": "R",
        # Humidity Ratio
        "w": "W", "omega": "W", "humrat": "W",
        # Enthalpy
        "hda": "Hda", "hha": "Hha",
        # Entropy
        "sda": "Sda", "sha": "Sha",
        # Specific Heat
        "cda": "Cda", "cpda": "Cda", 
        "cha": "Cha", "cpha": "Cha",
        # Conductivity and Viscosity
        "k": "K", "conductivity": "K",
        "mu": "MU", "viscosity": "MU",
        # Water mole fraction
        "psi_w": "Psi_w", "y": "Psi_w",
        # Specific Volume
        "vda": "Vda", "vha": "Vha",
        # Compressibility
        "z": "Z",
        # Density
        "dda": "Dda", "rhoda": "Dda",
        "dha": "Dha", "rhoha": "Dha",
    }
    return m.get(name.lower(), name)

def ConvertToSI_HA(name, value):
    if name in ("T", "Twb", "Tdp"):
        return value + 273.15
    if name in ("Hda", "Hha", "Sda", "Sha", "Cda", "Cha"):
        return value * 1000
    if name in ("P", "P_w"):
        return value * 1e5
    return value

def ConvertFromSI_HA(name, value):
    if name in ("T", "Twb", "Tdp"):
        return value - 273.15
    if name in ("Hda", "Hha", "Sda", "Sha", "Cda", "Cha"):
        return value / 1000
    if name in ("P", "P_w"):
        return value / 1e5
    return value

# === LibreOffice Calc Functions ===

def CProp_SI(output, name1, value1, name2, value2, fluid):
    """Calculate thermodynamic properties using SI units (no conversion)"""
    try:
        if not all([output, name1, name2, fluid]):
            return "Error: Missing parameter"
        if not isinstance(value1, (int, float)) or not isinstance(value2, (int, float)):
            return "Error: Non-numeric inputs"

        name1 = FormatName(name1)
        name2 = FormatName(name2)
        output = FormatName(output)

        result = PropsSI(output, name1, float(value1), name2, float(value2), fluid)

        if abs(result) >= 1.0E+308 or result != result:
            return f"Error: CoolProp failed. {get_global_param_string('errstring')}"
        return result
    except Exception as e:
        return f"Error: {str(e)}"

def CProp_E(output, name1, value1, name2, value2, fluid):
    """Calculate thermodynamic properties using engineering units (°C, bar, kJ/kg)"""
    try:
        if not all([output, name1, name2, fluid]):
            return "Error: Missing parameter"
        if not isinstance(value1, (int, float)) or not isinstance(value2, (int, float)):
            return "Error: Non-numeric inputs"

        name1 = FormatName(name1)
        name2 = FormatName(name2)
        output = FormatName(output)

        val1SI = ConvertToSI(name1, float(value1))
        val2SI = ConvertToSI(name2, float(value2))

        result = PropsSI(output, name1, val1SI, name2, val2SI, fluid)

        if abs(result) >= 1.0E+308 or result != result:
            return f"Error: CoolProp failed. {get_global_param_string('errstring')}"
        return ConvertFromSI(output, result)
    except Exception as e:
        return f"Error: {str(e)}"

def CProp(output, name1, value1, name2, value2, fluid):
    """Default function - uses engineering units. Alias for CProp_E."""
    return CProp_E(output, name1, value1, name2, value2, fluid)

def CPropHA_SI(output, name1, value1, name2, value2, name3, value3):
    """Calculate humid air properties using SI units (no conversion)"""
    try:
        if not all([output, name1, name2, name3]):
            return "Error: Missing parameter"
        if not all(isinstance(v, (int, float)) for v in [value1, value2, value3]):
            return "Error: Non-numeric inputs"

        name1 = FormatName_HA(name1)
        name2 = FormatName_HA(name2)
        name3 = FormatName_HA(name3)
        output = FormatName_HA(output)

        result = HAPropsSI(output, name1, float(value1), name2, float(value2), name3, float(value3))

        if abs(result) >= 1.0E+308 or result != result:
            return f"Error: CoolProp failed. {get_global_param_string('errstring')}"
        return result
    except Exception as e:
        return f"Error: {str(e)}"

def CPropHA_E(output, name1, value1, name2, value2, name3, value3):
    """Calculate humid air properties using engineering units (°C, bar, kJ/kg)"""
    try:
        if not all([output, name1, name2, name3]):
            return "Error: Missing parameter"
        if not all(isinstance(v, (int, float)) for v in [value1, value2, value3]):
            return "Error: Non-numeric inputs"

        name1 = FormatName_HA(name1)
        name2 = FormatName_HA(name2)
        name3 = FormatName_HA(name3)
        output = FormatName_HA(output)

        val1SI = ConvertToSI_HA(name1, float(value1))
        val2SI = ConvertToSI_HA(name2, float(value2))
        val3SI = ConvertToSI_HA(name3, float(value3))

        result = HAPropsSI(output, name1, val1SI, name2, val2SI, name3, val3SI)

        if abs(result) >= 1.0E+308 or result != result:
            return f"Error: CoolProp failed. {get_global_param_string('errstring')}"
        return ConvertFromSI_HA(output, result)
    except Exception as e:
        return f"Error: {str(e)}"

def CPropHA(output, name1, value1, name2, value2, name3, value3):
    """Default function - uses engineering units. Alias for CPropHA_E."""
    return CPropHA_E(output, name1, value1, name2, value2, name3, value3)


# ============================================================================
# LibreOffice Calc AddIn Implementation
# ============================================================================

class CoolPropAddIn(unohelper.Base, XAddIn, XServiceName, XServiceInfo):
    """
    LibreOffice Calc AddIn implementation for CoolProp functions.
    This class registers all functions with LibreOffice Calc.
    """
    
    def __init__(self, ctx):
        self.ctx = ctx
    
    # XServiceName methods
    def getServiceName(self):
        return IMPLE_NAME
    
    # XServiceInfo methods
    def getImplementationName(self):
        return IMPLE_NAME
    
    def supportsService(self, ServiceName):
        return ServiceName == SERVICE_NAME
    
    def getSupportedServiceNames(self):
        return (SERVICE_NAME,)
    
    # ========================================================================
    # Real Fluid Property Functions (SI Units)
    # ========================================================================
    
    def getCProp_SI(self, output, name1, value1, name2, value2, fluid):
        """Calculate thermodynamic properties using SI units (K, Pa, J/kg)"""
        return CProp_SI(output, name1, value1, name2, value2, fluid)
    
    # ========================================================================
    # Real Fluid Property Functions (Engineering Units)
    # ========================================================================
    
    def getCProp_E(self, output, name1, value1, name2, value2, fluid):
        """Calculate thermodynamic properties using engineering units (°C, bar, kJ/kg)"""
        return CProp_E(output, name1, value1, name2, value2, fluid)
    
    def getCProp(self, output, name1, value1, name2, value2, fluid):
        """Calculate thermodynamic properties (default: engineering units)"""
        return CProp(output, name1, value1, name2, value2, fluid)
    
    # ========================================================================
    # Humid Air Property Functions (SI Units)
    # ========================================================================
    
    def getCPropHA_SI(self, output, name1, value1, name2, value2, name3, value3):
        """Calculate humid air properties using SI units (K, Pa, J/kg)"""
        return CPropHA_SI(output, name1, value1, name2, value2, name3, value3)
    
    # ========================================================================
    # Humid Air Property Functions (Engineering Units)
    # ========================================================================
    
    def getCPropHA_E(self, output, name1, value1, name2, value2, name3, value3):
        """Calculate humid air properties using engineering units (°C, bar, kJ/kg)"""
        return CPropHA_E(output, name1, value1, name2, value2, name3, value3)
    
    def getCPropHA(self, output, name1, value1, name2, value2, name3, value3):
        """Calculate humid air properties (default: engineering units)"""
        return CPropHA(output, name1, value1, name2, value2, name3, value3)
    
    # ========================================================================
    # Function Descriptions for LibreOffice (for Function Wizard)
    # ========================================================================
    
    def getProgrammaticFuntionName(self, aDisplayName):
        """Map display names to internal function names"""
        function_map = {
            "CPROP_SI": "getCProp_SI",
            "CPROP_E": "getCProp_E",
            "CPROP": "getCProp",
            "CPROPHA_SI": "getCPropHA_SI",
            "CPROPHA_E": "getCPropHA_E",
            "CPROPHA": "getCPropHA",
        }
        return function_map.get(aDisplayName, "")
    
    def getDisplayFunctionName(self, aProgrammaticName):
        """Map internal function names to display names"""
        display_map = {
            "getCProp_SI": "CPROP_SI",
            "getCProp_E": "CPROP_E",
            "getCProp": "CPROP",
            "getCPropHA_SI": "CPROPHA_SI",
            "getCPropHA_E": "CPROPHA_E",
            "getCPropHA": "CPROPHA",
        }
        return display_map.get(aProgrammaticName, "")
    
    def getFunctionDescription(self, aProgrammaticName):
        """Provide function descriptions for the Function Wizard"""
        descriptions = {
            "getCProp_SI": "Calculate thermodynamic properties of real fluids using SI units (K, Pa, J/kg, etc.)",
            "getCProp_E": "Calculate thermodynamic properties of real fluids using engineering units (°C, bar, kJ/kg, etc.)",
            "getCProp": "Calculate thermodynamic properties of real fluids (default: engineering units)",
            "getCPropHA_SI": "Calculate humid air properties using SI units (K, Pa, J/kg, etc.)",
            "getCPropHA_E": "Calculate humid air properties using engineering units (°C, bar, kJ/kg, etc.)",
            "getCPropHA": "Calculate humid air properties (default: engineering units)",
        }
        return descriptions.get(aProgrammaticName, "")
    
    def getDisplayArgumentName(self, aProgrammaticFunctionName, nArgument):
        """Provide argument names for the Function Wizard"""
        arg_names = {
            "getCProp_SI": ["Output", "Name1", "Value1", "Name2", "Value2", "Fluid"],
            "getCProp_E": ["Output", "Name1", "Value1", "Name2", "Value2", "Fluid"],
            "getCProp": ["Output", "Name1", "Value1", "Name2", "Value2", "Fluid"],
            "getCPropHA_SI": ["Output", "Name1", "Value1", "Name2", "Value2", "Name3", "Value3"],
            "getCPropHA_E": ["Output", "Name1", "Value1", "Name2", "Value2", "Name3", "Value3"],
            "getCPropHA": ["Output", "Name1", "Value1", "Name2", "Value2", "Name3", "Value3"],
        }
        args = arg_names.get(aProgrammaticFunctionName, [])
        return args[nArgument] if nArgument < len(args) else ""
    
    def getArgumentDescription(self, aProgrammaticFunctionName, nArgument):
        """Provide argument descriptions for the Function Wizard"""
        arg_descriptions = {
            "getCProp_SI": [
                "Property to calculate (e.g., 'H', 'S', 'T', 'P')",
                "First property name (e.g., 'T', 'P', 'H')",
                "First property value (in SI units)",
                "Second property name",
                "Second property value (in SI units)",
                "Fluid name (e.g., 'Water', 'R134a')"
            ],
            "getCProp_E": [
                "Property to calculate (e.g., 'H', 'S', 'T', 'P')",
                "First property name (e.g., 'T', 'P', 'H')",
                "First property value (in engineering units)",
                "Second property name",
                "Second property value (in engineering units)",
                "Fluid name (e.g., 'Water', 'R134a')"
            ],
            "getCProp": [
                "Property to calculate (e.g., 'H', 'S', 'T', 'P')",
                "First property name (e.g., 'T', 'P', 'H')",
                "First property value (in engineering units)",
                "Second property name",
                "Second property value (in engineering units)",
                "Fluid name (e.g., 'Water', 'R134a')"
            ],
            "getCPropHA_SI": [
                "Property to calculate (e.g., 'W', 'Hha', 'RH')",
                "First property name (e.g., 'T', 'P')",
                "First property value (in SI units)",
                "Second property name",
                "Second property value (in SI units)",
                "Third property name (e.g., 'R' for relative humidity)",
                "Third property value (in SI units)"
            ],
            "getCPropHA_E": [
                "Property to calculate (e.g., 'W', 'Hha', 'RH')",
                "First property name (e.g., 'T', 'P')",
                "First property value (in engineering units)",
                "Second property name",
                "Second property value (in engineering units)",
                "Third property name (e.g., 'RH' for relative humidity)",
                "Third property value (in engineering units)"
            ],
            "getCPropHA": [
                "Property to calculate (e.g., 'W', 'Hha', 'RH')",
                "First property name (e.g., 'T', 'P')",
                "First property value (in engineering units)",
                "Second property name",
                "Second property value (in engineering units)",
                "Third property name (e.g., 'RH' for relative humidity)",
                "Third property value (in engineering units)"
            ],
        }
        descriptions = arg_descriptions.get(aProgrammaticFunctionName, [])
        return descriptions[nArgument] if nArgument < len(descriptions) else ""
    
    def getCategoryDisplayName(self, aProgrammaticFunctionName):
        """Assign functions to a category in LibreOffice"""
        return "CoolProp"


# ============================================================================
# Service Registration
# ============================================================================

def createInstance(ctx):
    """Factory function to create an instance of the AddIn"""
    return CoolPropAddIn(ctx)


# g_ImplementationHelper is used for component registration
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    createInstance,
    IMPLE_NAME,
    (SERVICE_NAME,),
)
