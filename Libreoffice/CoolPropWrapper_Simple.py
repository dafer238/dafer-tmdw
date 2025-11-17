"""
CoolProp Wrapper for LibreOffice Calc - Simplified Version
Provides thermodynamic property calculations using CoolProp library
"""

try:
    from CoolProp.CoolProp import PropsSI, HAPropsSI, get_global_param_string
    COOLPROP_AVAILABLE = True
except ImportError:
    COOLPROP_AVAILABLE = False
    def PropsSI(*args):
        return "ERROR: CoolProp not installed"
    def HAPropsSI(*args):
        return "ERROR: CoolProp not installed"


def ConvertToSI(name, value):
    """Convert from engineering units to SI units"""
    conversions = {
        "T": lambda x: x + 273.15,  # °C to K
        "P": lambda x: x * 1e5,  # bar to Pa
        "H": lambda x: x * 1000,  # kJ/kg to J/kg
        "U": lambda x: x * 1000,
        "S": lambda x: x * 1000,
        "Cp": lambda x: x * 1000,
        "Cpmass": lambda x: x * 1000,
        "Cvmass": lambda x: x * 1000,
    }
    return conversions.get(name, lambda x: x)(value)


def ConvertFromSI(name, value):
    """Convert from SI units to engineering units"""
    conversions = {
        "T": lambda x: x - 273.15,  # K to °C
        "P": lambda x: x / 1e5,  # Pa to bar
        "H": lambda x: x / 1000,  # J/kg to kJ/kg
        "U": lambda x: x / 1000,
        "S": lambda x: x / 1000,
        "Cp": lambda x: x / 1000,
        "Cpmass": lambda x: x / 1000,
        "Cvmass": lambda x: x / 1000,
    }
    return conversions.get(name, lambda x: x)(value)


def FormatName(name):
    """Normalize property name aliases to CoolProp standard names"""
    if not isinstance(name, str):
        return str(name)
    
    name_lower = name.lower()
    aliases = {
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
        # Density
        "rho": "D", "dens": "D", "dmass": "D",
        # Specific Heat
        "cvmass": "Cvmass", "cv": "Cvmass",
        "cpmass": "Cpmass", "cp": "Cpmass",
        # Quality
        "q": "Q", "quality": "Q", "x": "Q",
        # Other properties
        "mu": "MU", "viscosity": "MU",
        "k": "K", "conductivity": "K",
        "z": "Z",
        "mm": "MM", "molar_mass": "MM",
    }
    return aliases.get(name_lower, name)


def ConvertToSI_HA(name, value):
    """Convert humid air properties from engineering to SI units"""
    if name in ("T", "Twb", "Tdp"):
        return value + 273.15  # °C to K
    if name in ("Hda", "Hha", "Sda", "Sha", "Cda", "Cha"):
        return value * 1000  # kJ/kg to J/kg
    if name in ("P", "P_w"):
        return value * 1e5  # bar to Pa
    return value


def ConvertFromSI_HA(name, value):
    """Convert humid air properties from SI to engineering units"""
    if name in ("T", "Twb", "Tdp"):
        return value - 273.15  # K to °C
    if name in ("Hda", "Hha", "Sda", "Sha", "Cda", "Cha"):
        return value / 1000  # J/kg to kJ/kg
    if name in ("P", "P_w"):
        return value / 1e5  # Pa to bar
    return value


def FormatName_HA(name):
    """Normalize humid air property name aliases"""
    if not isinstance(name, str):
        return str(name)
    
    name_lower = name.lower()
    aliases = {
        "twb": "Twb", "wetbulb": "Twb", "t_wb": "Twb",
        "tdp": "Tdp", "dewpoint": "Tdp", "t_dp": "Tdp",
        "t": "T", "tdb": "T", "t_db": "T",
        "p": "P", "p_w": "P_w",
        "r": "R", "rh": "R", "relhum": "R",
        "w": "W", "omega": "W", "humrat": "W",
        "hda": "Hda", "hha": "Hha",
        "sda": "Sda", "sha": "Sha",
        "cda": "Cda", "cpda": "Cda",
        "cha": "Cha", "cpha": "Cha",
    }
    return aliases.get(name_lower, name)


# ============================================================================
# Main calculation functions (called directly from Calc)
# ============================================================================

def CPROP_SI(output, name1, value1, name2, value2, fluid):
    """Calculate thermodynamic properties using SI units (K, Pa, J/kg)"""
    try:
        if not COOLPROP_AVAILABLE:
            return "ERROR: CoolProp not installed. See README_LibreOffice.md"
        
        # Convert to strings and floats
        output = str(output)
        name1 = FormatName(str(name1))
        name2 = FormatName(str(name2))
        output = FormatName(output)
        fluid = str(fluid)
        value1 = float(value1)
        value2 = float(value2)
        
        result = PropsSI(output, name1, value1, name2, value2, fluid)
        
        if not isinstance(result, (int, float)) or result >= 1.0E+308 or result <= -1.0E+308:
            return f"ERROR: CoolProp calculation failed"
        
        return result
    except Exception as e:
        return f"ERROR: {str(e)}"


def CPROP_E(output, name1, value1, name2, value2, fluid):
    """Calculate thermodynamic properties using engineering units (°C, bar, kJ/kg)"""
    try:
        if not COOLPROP_AVAILABLE:
            return "ERROR: CoolProp not installed. See README_LibreOffice.md"
        
        # Convert to strings and floats
        output = str(output)
        name1 = FormatName(str(name1))
        name2 = FormatName(str(name2))
        output = FormatName(output)
        fluid = str(fluid)
        value1 = float(value1)
        value2 = float(value2)
        
        # Convert inputs to SI
        val1SI = ConvertToSI(name1, value1)
        val2SI = ConvertToSI(name2, value2)
        
        result = PropsSI(output, name1, val1SI, name2, val2SI, fluid)
        
        if not isinstance(result, (int, float)) or result >= 1.0E+308 or result <= -1.0E+308:
            return f"ERROR: CoolProp calculation failed"
        
        # Convert output from SI
        return ConvertFromSI(output, result)
    except Exception as e:
        return f"ERROR: {str(e)}"


def CPROP(output, name1, value1, name2, value2, fluid):
    """Calculate thermodynamic properties (default: engineering units)"""
    return CPROP_E(output, name1, value1, name2, value2, fluid)


def CPROPHA_SI(output, name1, value1, name2, value2, name3, value3):
    """Calculate humid air properties using SI units (K, Pa, J/kg)"""
    try:
        if not COOLPROP_AVAILABLE:
            return "ERROR: CoolProp not installed. See README_LibreOffice.md"
        
        # Convert to strings and floats
        output = FormatName_HA(str(output))
        name1 = FormatName_HA(str(name1))
        name2 = FormatName_HA(str(name2))
        name3 = FormatName_HA(str(name3))
        value1 = float(value1)
        value2 = float(value2)
        value3 = float(value3)
        
        result = HAPropsSI(output, name1, value1, name2, value2, name3, value3)
        
        if not isinstance(result, (int, float)) or result >= 1.0E+308 or result <= -1.0E+308:
            return f"ERROR: CoolProp calculation failed"
        
        return result
    except Exception as e:
        return f"ERROR: {str(e)}"


def CPROPHA_E(output, name1, value1, name2, value2, name3, value3):
    """Calculate humid air properties using engineering units (°C, bar, kJ/kg)"""
    try:
        if not COOLPROP_AVAILABLE:
            return "ERROR: CoolProp not installed. See README_LibreOffice.md"
        
        # Convert to strings and floats
        output = FormatName_HA(str(output))
        name1 = FormatName_HA(str(name1))
        name2 = FormatName_HA(str(name2))
        name3 = FormatName_HA(str(name3))
        value1 = float(value1)
        value2 = float(value2)
        value3 = float(value3)
        
        # Convert inputs to SI
        val1SI = ConvertToSI_HA(name1, value1)
        val2SI = ConvertToSI_HA(name2, value2)
        val3SI = ConvertToSI_HA(name3, value3)
        
        result = HAPropsSI(output, name1, val1SI, name2, val2SI, name3, val3SI)
        
        if not isinstance(result, (int, float)) or result >= 1.0E+308 or result <= -1.0E+308:
            return f"ERROR: CoolProp calculation failed"
        
        # Convert output from SI
        return ConvertFromSI_HA(output, result)
    except Exception as e:
        return f"ERROR: {str(e)}"


def CPROPHA(output, name1, value1, name2, value2, name3, value3):
    """Calculate humid air properties (default: engineering units)"""
    return CPROPHA_E(output, name1, value1, name2, value2, name3, value3)


