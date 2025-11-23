using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.Integration;

public static class CoolPropWrapper
{
    // Windows API for LoadLibrary and SetDllDirectory
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr LoadLibrary(string dllToLoad);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GetProcAddress(IntPtr hModule, string procedureName);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool SetDllDirectory(string lpPathName);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool AddDllDirectory(string NewDirectory);

    // Static constructor to load CoolProp.dll
    static CoolPropWrapper()
    {
        LoadCoolPropDll();
    }

    private static void LoadCoolPropDll()
    {
        try
        {
            // Multiple approaches to find and load CoolProp.dll
            string[] searchPaths = {
                // Excel-DNA specific: XLL directory
                Path.GetDirectoryName(ExcelDnaUtil.XllPath),
                // Current assembly location
                Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                // Current directory
                Directory.GetCurrentDirectory()
            };

            foreach (string searchPath in searchPaths)
            {
                if (string.IsNullOrEmpty(searchPath)) continue;

                string coolPropPath = Path.Combine(searchPath, "CoolProp.dll");
                if (File.Exists(coolPropPath))
                {
                    // Try to set the DLL directory first
                    SetDllDirectory(searchPath);
                    
                    // Try to load with full path
                    IntPtr handle = LoadLibrary(coolPropPath);
                    if (handle != IntPtr.Zero)
                    {
                        // Successfully loaded, we're done
                        return;
                    }
                }
            }

            // Final fallback - let Windows search
            LoadLibrary("CoolProp.dll");
        }
        catch
        {
            // If all else fails, let the DllImport handle it during first call
        }
    }

    // Diagnostic function to help troubleshoot DLL loading
    [ExcelFunction(Description = "Diagnostic function to check CoolProp DLL loading paths")]
    public static object CPropDiag()
    {
        try
        {
            string[] searchPaths = GetValidSearchPaths();

            string result = "CoolProp.dll search paths:\n";
            for (int i = 0; i < searchPaths.Length; i++)
            {
                try
                {
                    string coolPropPath = Path.Combine(searchPaths[i], "CoolProp.dll");
                    bool exists = File.Exists(coolPropPath);
                    result += $"{i + 1}. {searchPaths[i]} - {(exists ? "FOUND" : "NOT FOUND")}\n";
                }
                catch (Exception pathEx)
                {
                    result += $"{i + 1}. {searchPaths[i]} - ERROR: {pathEx.Message}\n";
                }
            }

            // Additional debug info
            result += "\nDebug info:\n";
            try
            {
                result += $"XllPath: {ExcelDnaUtil.XllPath ?? "NULL"}\n";
            }
            catch (Exception ex)
            {
                result += $"XllPath: ERROR - {ex.Message}\n";
            }

            try
            {
                result += $"Assembly Location: {Assembly.GetExecutingAssembly().Location ?? "NULL"}\n";
            }
            catch (Exception ex)
            {
                result += $"Assembly Location: ERROR - {ex.Message}\n";
            }

            try
            {
                result += $"Current Directory: {Directory.GetCurrentDirectory() ?? "NULL"}\n";
            }
            catch (Exception ex)
            {
                result += $"Current Directory: ERROR - {ex.Message}\n";
            }
            
            return result;
        }
        catch (Exception ex)
        {
            return $"Diagnostic error: {ex.Message}";
        }
    }

    // Dynamic loading approach - store the loaded handle
    private static IntPtr _coolPropHandle = IntPtr.Zero;

    // Delegate types for CoolProp functions
    private delegate double PropsSI_Delegate(string output, string name1, double value1, string name2, double value2, string fluid);
    private delegate double HAPropsSI_Delegate(string output, string name1, double value1, string name2, double value2, string name3, double value3);
    private delegate IntPtr get_global_param_string_Delegate(string param);

    // Static function pointers
    private static PropsSI_Delegate _PropsSI;
    private static HAPropsSI_Delegate _HAPropsSI;
    private static get_global_param_string_Delegate _get_global_param_string;

    // Wrapper functions that use dynamic loading
    public static double PropsSI(string output, string name1, double value1, string name2, double value2, string fluid)
    {
        EnsureCoolPropLoaded();
        return _PropsSI(output, name1, value1, name2, value2, fluid);
    }

    public static double HAPropsSI(string output, string name1, double value1, string name2, double value2, string name3, double value3)
    {
        EnsureCoolPropLoaded();
        return _HAPropsSI(output, name1, value1, name2, value2, name3, value3);
    }

    private static IntPtr get_global_param_string(string param)
    {
        EnsureCoolPropLoaded();
        return _get_global_param_string(param);
    }

    private static void EnsureCoolPropLoaded()
    {
        if (_coolPropHandle != IntPtr.Zero && _PropsSI != null) return;

        // Find and load CoolProp.dll with safe path handling
        string[] searchPaths = GetValidSearchPaths();

        foreach (string searchPath in searchPaths)
        {
            if (string.IsNullOrEmpty(searchPath)) continue;

            try
            {
                string coolPropPath = Path.Combine(searchPath, "CoolProp.dll");
                if (File.Exists(coolPropPath))
                {
                    _coolPropHandle = LoadLibrary(coolPropPath);
                    if (_coolPropHandle != IntPtr.Zero)
                    {
                        LoadFunctionPointers();
                        return;
                    }
                }
            }
            catch (ArgumentException)
            {
                // Invalid path format, skip this path
                continue;
            }
            catch (PathTooLongException)
            {
                // Path too long, skip this path
                continue;
            }
        }

        // Final attempt with system search
        _coolPropHandle = LoadLibrary("CoolProp.dll");
        if (_coolPropHandle != IntPtr.Zero)
        {
            LoadFunctionPointers();
            return;
        }

        throw new DllNotFoundException($"CoolProp.dll could not be loaded from any valid search path. Paths tried: {string.Join("; ", searchPaths)}");
    }

    private static string[] GetValidSearchPaths()
    {
        var validPaths = new List<string>();

        // Try Excel-DNA specific path
        try
        {
            string xllPath = ExcelDnaUtil.XllPath;
            if (!string.IsNullOrEmpty(xllPath))
            {
                string xllDir = Path.GetDirectoryName(xllPath);
                if (!string.IsNullOrEmpty(xllDir) && Directory.Exists(xllDir))
                    validPaths.Add(xllDir);
            }
        }
        catch { /* Ignore errors getting XLL path */ }

        // Try assembly location
        try
        {
            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            if (!string.IsNullOrEmpty(assemblyPath))
            {
                string assemblyDir = Path.GetDirectoryName(assemblyPath);
                if (!string.IsNullOrEmpty(assemblyDir) && Directory.Exists(assemblyDir))
                    validPaths.Add(assemblyDir);
            }
        }
        catch { /* Ignore errors getting assembly location */ }

        // Try current directory
        try
        {
            string currentDir = Directory.GetCurrentDirectory();
            if (!string.IsNullOrEmpty(currentDir) && Directory.Exists(currentDir))
                validPaths.Add(currentDir);
        }
        catch { /* Ignore errors getting current directory */ }

        return validPaths.ToArray();
    }

    private static void LoadFunctionPointers()
    {
        IntPtr propsSIPtr = GetProcAddress(_coolPropHandle, "PropsSI");
        IntPtr haPropsSIPtr = GetProcAddress(_coolPropHandle, "HAPropsSI");
        IntPtr getErrorPtr = GetProcAddress(_coolPropHandle, "get_global_param_string");

        if (propsSIPtr != IntPtr.Zero)
            _PropsSI = Marshal.GetDelegateForFunctionPointer<PropsSI_Delegate>(propsSIPtr);
        if (haPropsSIPtr != IntPtr.Zero)
            _HAPropsSI = Marshal.GetDelegateForFunctionPointer<HAPropsSI_Delegate>(haPropsSIPtr);
        if (getErrorPtr != IntPtr.Zero)
            _get_global_param_string = Marshal.GetDelegateForFunctionPointer<get_global_param_string_Delegate>(getErrorPtr);

        if (_PropsSI == null || _HAPropsSI == null || _get_global_param_string == null)
        {
            throw new EntryPointNotFoundException("Required functions not found in CoolProp.dll");
        }
    }

    // Function to retrieve the error message from CoolProp
    private static string GetCoolPropError()
    {
        IntPtr ptr = get_global_param_string("errstring");
        return Marshal.PtrToStringAnsi(ptr) ?? "Unknown error";
    }

    // Function to calculate thermodynamic properties using SI units (no conversion)
    [ExcelFunction(Description = "Calculate thermodynamic properties of real fluids using CoolProp with SI units (K, Pa, J/kg, etc.).")]
    public static object CProp_SI(string output, string name1, object value1, string name2, object value2, string fluid)
    {
        // Check for missing or invalid inputs
        if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
        if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
        if (value1 == null || !(value1 is double)) return "Error: First property value is missing or not a number.";
        if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
        if (value2 == null || !(value2 is double)) return "Error: Second property value is missing or not a number.";
        if (string.IsNullOrWhiteSpace(fluid)) return "Error: Fluid name is missing.";

        // Normalize parameter names
        name1 = FormatName(name1);
        name2 = FormatName(name2);
        output = FormatName(output);

        // Call CoolProp directly with SI units (no conversion)
        double result;
        try
        {
            result = PropsSI(output, name1, (double)value1, name2, (double)value2, fluid);

            // Check if the result is a large number (potential error)
            if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
            {
                return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
            }
        }
        catch (DllNotFoundException ex)
        {
            return $"Error: CoolProp.dll not found in any search path. Use CPropDiag() function to see paths checked. {ex.Message}";
        }
        catch (EntryPointNotFoundException ex)
        {
            return $"Error: Required functions not found in CoolProp.dll. Check DLL version and architecture match. {ex.Message}";
        }
        catch (Exception ex)
        {
            return $"Error: Exception occurred while computing property. {ex.Message}";
        }

        // Return result in SI units (no conversion)
        return result;
    }

    // Function to calculate thermodynamic properties using engineering units
    [ExcelFunction(Description = "Calculate thermodynamic properties of real fluids using CoolProp with engineering units (°C, bar, kJ/kg, etc.).")]
    public static object CProp_E(string output, string name1, object value1, string name2, object value2, string fluid)
    {
        // Check for missing or invalid inputs
        if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
        if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
        if (value1 == null || !(value1 is double)) return "Error: First property value is missing or not a number.";
        if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
        if (value2 == null || !(value2 is double)) return "Error: Second property value is missing or not a number.";
        if (string.IsNullOrWhiteSpace(fluid)) return "Error: Fluid name is missing.";

        // Normalize parameter names
        name1 = FormatName(name1);
        name2 = FormatName(name2);
        output = FormatName(output);

        // Convert inputs to SI units
        double val1SI = ConvertToSI(name1, (double)value1);
        double val2SI = ConvertToSI(name2, (double)value2);

        // Call CoolProp for the requested property
        double result;
        try
        {
            result = PropsSI(output, name1, val1SI, name2, val2SI, fluid);

            // Check if the result is a large number (potential error)
            if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
            {
                return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
            }
        }
        catch (DllNotFoundException ex)
        {
            return $"Error: CoolProp.dll not found in any search path. Use CPropDiag() function to see paths checked. {ex.Message}";
        }
        catch (EntryPointNotFoundException ex)
        {
            return $"Error: Required functions not found in CoolProp.dll. Check DLL version and architecture match. {ex.Message}";
        }
        catch (Exception ex)
        {
            return $"Error: Exception occurred while computing property. {ex.Message}";
        }

        // Convert output to engineering units
        return ConvertFromSI(output, result);
    }

    // Default function (uses engineering units)
    [ExcelFunction(Description = "Calculate thermodynamic properties of real fluids using CoolProp with engineering units (°C, bar, kJ/kg, etc.). Alias for CProp_E.")]
    public static object CProp(string output, string name1, object value1, string name2, object value2, string fluid)
    {
        return CProp_E(output, name1, value1, name2, value2, fluid);
    }


[ExcelFunction(Description = "Calculate thermodynamic properties of humid air using CoolProp with SI units (K, Pa, J/kg, etc.).")]
public static object CPropHA_SI(string output, string name1, object value1, string name2, object value2, string name3, object value3)
{
    // Check for missing or invalid inputs
    if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
    if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
    if (value1 == null || !(value1 is double)) return "Error: First property value is missing or not a number.";
    if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
    if (value2 == null || !(value2 is double)) return "Error: Second property value is missing or not a number.";
    if (string.IsNullOrWhiteSpace(name3)) return "Error: Third property name is missing.";
    if (value3 == null || !(value3 is double)) return "Error: Third property value is missing or not a number.";

    // Normalize parameter names
    name1 = FormatName_HA(name1);
    name2 = FormatName_HA(name2);
    name3 = FormatName_HA(name3);
    output = FormatName_HA(output);

    // Call CoolProp directly with SI units (no conversion)
    double result;
    try
    {
        result = HAPropsSI(output, name1, (double)value1, name2, (double)value2, name3, (double)value3);

        // Check if the result is a large number (potential error)
        if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
        {
            return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
        }
    }
    catch (DllNotFoundException ex)
    {
        return $"Error: CoolProp.dll not found in any search path. Use CPropDiag() function to see paths checked. {ex.Message}";
    }
    catch (EntryPointNotFoundException ex)
    {
        return $"Error: Required functions not found in CoolProp.dll. Check DLL version and architecture match. {ex.Message}";
    }
    catch (Exception ex)
    {
        return $"Error: Exception occurred while computing property. {ex.Message}";
    }

    // Return result in SI units (no conversion)
    return result;
}

[ExcelFunction(Description = "Calculate thermodynamic properties of humid air using CoolProp with engineering units (°C, bar, kJ/kg, etc.).")]
public static object CPropHA_E(string output, string name1, object value1, string name2, object value2, string name3, object value3)
{
    // Check for missing or invalid inputs
    if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
    if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
    if (value1 == null || !(value1 is double)) return "Error: First property value is missing or not a number.";
    if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
    if (value2 == null || !(value2 is double)) return "Error: Second property value is missing or not a number.";
    if (string.IsNullOrWhiteSpace(name3)) return "Error: Third property name is missing.";
    if (value3 == null || !(value3 is double)) return "Error: Third property value is missing or not a number.";

    // Normalize parameter names
    name1 = FormatName_HA(name1);
    name2 = FormatName_HA(name2);
    name3 = FormatName_HA(name3);
    output = FormatName_HA(output);

    // Convert inputs to SI units
    double val1SI = ConvertToSI_HA(name1, (double)value1);
    double val2SI = ConvertToSI_HA(name2, (double)value2);
    double val3SI = ConvertToSI_HA(name3, (double)value3);

    // Call CoolProp for the requested property
    double result;
    try
    {
        result = HAPropsSI(output, name1, val1SI, name2, val2SI, name3, val3SI);

        // Check if the result is a large number (potential error)
        if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
        {
            return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
        }
    }
    catch (DllNotFoundException ex)
    {
        return $"Error: CoolProp.dll not found in any search path. Use CPropDiag() function to see paths checked. {ex.Message}";
    }
    catch (EntryPointNotFoundException ex)
    {
        return $"Error: Required functions not found in CoolProp.dll. Check DLL version and architecture match. {ex.Message}";
    }
    catch (Exception ex)
    {
        return $"Error: Exception occurred while computing property. {ex.Message}";
    }

    // Convert output to engineering units
    return ConvertFromSI_HA(output, result);
    }

    // Default function (uses engineering units)
    [ExcelFunction(Description = "Calculate thermodynamic properties of humid air using CoolProp with engineering units (°C, bar, kJ/kg, etc.). Alias for CPropHA_E.")]
    public static object CPropHA(string output, string name1, object value1, string name2, object value2, string name3, object value3)
    {
        return CPropHA_E(output, name1, value1, name2, value2, name3, value3);
    }


    // Normalize name capitalization to match the expected format
    private static string FormatName(string name)
    {
        switch (name.ToLower())
        {
            case "t":
            case "temp":
            case "temperature":
                return "T";
            case "p":
            case "pres":
            case "pressure":
                return "P";
            case "h":
            case "enth":
            case "enthalpy":
            case "hmass":
                return "H";
            case "u":
            case "internalenergy":
            case "umass":
                return "U";
            case "s":
            case "entr":
            case "entropy":
            case "smass":
                return "S";
            case "dmolar":
            case "dmol":
                return "Dmolar";
            case "delta":
                return "Delta";
            case "rho":
            case "dens":
            case "dmass":
                return "D";
            case "cvmass":
            case "cv":
                return "Cvmass";
            case "cpmass":
            case "cp":
                return "Cpmass";
            case "cpmolar":
            case "cpmol":
                return "Cpmolar";
            case "cvmolar":
            case "cvmol":
                return "Cvmolar";
            case "q":
            case "quality":
            case "x":
                return "Q";
            case "tau":
                return "Tau";
            case "alpha0":
                return "Alpha0";
            case "alphar":
                return "Alphar";
            case "speed_of_sound":
            case "a":
                return "A";
            case "bvirial":
                return "Bvirial";
            case "k":
            case "conductivity":
                return "K";
            case "cvirial":
                return "Cvirial";
            case "dipoe_moment":
                return "DIPOLE_MOMENT";
            case "fh":
                return "FH";
            case "g":
            case "gmass":
                return "G";
            case "helmoltzmass":
                return "HELMHOLTZMASS";
            case "helmholtzmolar":
                return "HELMHOLTZMOLAR";
            case "gamma":
                return "gamma";
            case "isobaric_expansion_coefficient":
                return "isobaric_expansion_coefficient";
            case "isothermal_compressibility":
                return "isothermal_compressibility";
            case "surface_tension":
                return "surface_tension";
            case "mm":
            case "molar_mass":
                return "MM";
            case "pcrit":
            case "p_critical":
                return "Pcrit";
            case "phase":
                return "Phase";
            case "pmax":
                return "pmax";
            case "pmin":
                return "pmin";
            case "prandtl":
                return "Prandtl";
            case "ptriple":
                return "ptriple";
            case "p_reducing":
                return "p_reducing";
            case "rhocrit":
                return "rhocrit";
            case "rhomass_reducing":
                return "rhomass_reducing";
            case "smolar_residual":
                return "Smolar_residual";
            case "tcrit":
                return "Tcrit";
            case "tmax":
                return "Tmax";
            case "tmin":
                return "Tmin";
            case "ttriple":
                return "Ttriple";
            case "t_freeze":
                return "T_freeze";
            case "t_reducing":
                return "T_reducing";
            case "mu":
            case "viscosity":
                return "MU";
            case "z":
                return "Z";
            default:
                return name; // Return as is if no match is found
        }
    }

    // Convert from custom units to SI units
    private static double ConvertToSI(string name, double value)
    {
        switch (name)
        {
            case "T": // Temperature (°C to K)
                return value + 273.15;
            case "P": // Pressure (bar to Pa)
                return value * 1e5;
            case "H": // Specific Enthalpy (kJ/kg to J/kg)
            case "U": // Specific Internal Energy (kJ/kg to J/kg)
            case "S": // Specific Entropy (kJ/kg to J/kg)
            case "Cp": // Mass specific heat at constant pressure (kJ/kg/K to J/kg/K)
            case "Cpmass":
            case "Cvmass": // Mass specific heat at constant volume (kJ/kg/K to J/kg/K)
                return value * 1000;
            default:
                return value; // No conversion if unit not found
        }
    }

    // Convert from SI units to custom units
    private static double ConvertFromSI(string name, double value)
    {
        switch (name)
        {
            case "T": // Temperature (K to °C)
                return value - 273.15;
            case "P": // Pressure (Pa to bar)
                return value / 1e5;
            case "H": // Specific Enthalpy (J/kg to kJ/kg)
            case "U": // Specific Internal Energy (J/kg to kJ/kg)
            case "S": // Specific Entropy (J/kg to kJ/kg)
            case "Cp": // Mass specific heat at constant pressure (J/kg/K to kJ/kg/K)
            case "Cpmass":
            case "Cvmass": // Mass specific heat at constant volume (J/kg/K to kJ/kg/K)
                return value / 1000;
            default:
                return value; // No conversion if unit not found
        }
    }

    // Normalize name capitalization for HA properties
    private static string FormatName_HA(string name)
    {
        switch (name.ToLower())
        {
            case "twb":
            case "t_wb":
            case "wetbulb":
                return "Twb";
            case "cda":
            case "cpda":
                return "Cda";
            case "cha":
            case "cpha":
                return "Cha";
            case "tdp":
            case "dewpoint":
            case "t_dp":
                return "Tdp";
            case "hda":
                return "Hda";
            case "hha":
                return "Hha";
            case "k":
            case "conductivity":
                return "K";
            case "mu":
            case "viscosity":
                return "MU";
            case "psi_w":
            case "y":
                return "Psi_w";
            case "p":
                return "P";
            case "p_w":
                return "P_w";
            case "r":
            case "rh":
            case "relhum":
                return "R";
            case "sda":
                return "Sda";
            case "sha":
                return "Sha";
            case "t":
            case "tdb":
            case "t_db":
                return "T";
            case "vda":
                return "Vda";
            case "vha":
                return "Vha";
            case "w":
            case "omega":
            case "humrat":
                return "W";
            case "z":
                return "Z";
            case "dda":
            case "rhoda":
                return "Dda";
            case "dha":
            case "rhoha":
                return "Dha";
            default:
                return name; // Return as is if no match is found
        }
    }

    // Convert from custom units to SI units for HA properties
    private static double ConvertToSI_HA(string name, double value)
    {
        switch (name)
        {
            case "Twb": // Wet Bulb Temperature (°C to K)
            case "Tdp": // Dew Point Temperature (°C to K)
            case "T": // Dry Bulb Temperature (°C to K)
                return value + 273.15;
            case "Cda": // Specific heat of dry air (kJ/kg/K to J/kg/K)
            case "Cha": // Specific heat of humid air (kJ/kg/K to J/kg/K)
            case "Hda": // Specific Enthalpy of dry air (kJ/kg to J/kg)
            case "Hha": // Specific Enthalpy of humid air (kJ/kg to J/kg)
            case "Sda": // Specific entropy of dry air (kJ/kg/K to J/kg/K)
            case "Sha": // Specific entropy of humid air (kJ/kg/K to J/kg/K)
                return value * 1000;
            case "P": // Pressure (bar to Pa)
            case "P_w": // Partial pressure of water (bar to Pa)
                return value * 1e5;
            default:
                return value; // No conversion if unit not found
        }
    }

    // Convert from SI units to custom units for HA properties
    private static double ConvertFromSI_HA(string name, double value)
    {
        switch (name)
        {
            case "Twb": // Wet Bulb Temperature (K to °C)
            case "Tdp": // Dew Point Temperature (K to °C)
            case "T": // Dry Bulb Temperature (K to °C)
                return value - 273.15;
            case "Cda": // Specific heat of dry air (J/kg/K to kJ/kg/K)
            case "Cha": // Specific heat of humid air (J/kg/K to kJ/kg/K)
            case "Hda": // Specific Enthalpy of dry air (J/kg to kJ/kg)
            case "Hha": // Specific Enthalpy of humid air (J/kg to kJ/kg)
            case "Sda": // Specific entropy of dry air (J/kg/K to kJ/kg/K)
            case "Sha": // Specific entropy of humid air (J/kg/K to kJ/kg/K)
                return value / 1000;
            case "P": // Pressure (Pa to bar)
            case "P_w": // Partial pressure of water (Pa to bar)
                return value / 1e5;
            default:
                return value; // No conversion if unit not found
        }
    }

}

