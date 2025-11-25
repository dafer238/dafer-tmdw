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
    [ExcelFunction(Name = "CPropDiag", Description = "Diagnostic function to check CoolProp DLL loading paths")]
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
    private delegate double Props1SI_Delegate(string output, string fluid);
    private delegate long PhaseSI_Delegate(string name1, double value1, string name2, double value2, string fluid, IntPtr output, int n);
    private delegate double HAPropsSI_Delegate(string output, string name1, double value1, string name2, double value2, string name3, double value3);
    private delegate long get_global_param_string_Delegate(string param, IntPtr output, int n);
    private delegate long get_fluid_param_string_Delegate(string fluid, string param, IntPtr output, int n);

    // Static function pointers
    private static PropsSI_Delegate _PropsSI;
    private static Props1SI_Delegate _Props1SI;
    private static PhaseSI_Delegate _PhaseSI;
    private static HAPropsSI_Delegate _HAPropsSI;
    private static get_global_param_string_Delegate _get_global_param_string;
    private static get_fluid_param_string_Delegate _get_fluid_param_string;

    // Internal wrapper functions that use dynamic loading
    private static double PropsSI_Internal(string output, string name1, double value1, string name2, double value2, string fluid)
    {
        EnsureCoolPropLoaded();
        return _PropsSI(output, name1, value1, name2, value2, fluid);
    }

    private static double Props1SI_Internal(string output, string fluid)
    {
        EnsureCoolPropLoaded();
        return _Props1SI(output, fluid);
    }

    private static string PhaseSI_Internal(string name1, double value1, string name2, double value2, string fluid)
    {
        EnsureCoolPropLoaded();
        IntPtr buffer = Marshal.AllocHGlobal(2000);
        try
        {
            long result = _PhaseSI(name1, value1, name2, value2, fluid, buffer, 2000);
            if (result == 0)
                return GetCoolPropError();
            return Marshal.PtrToStringAnsi(buffer);
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    private static double HAPropsSI_Internal(string output, string name1, double value1, string name2, double value2, string name3, double value3)
    {
        EnsureCoolPropLoaded();
        return _HAPropsSI(output, name1, value1, name2, value2, name3, value3);
    }

    private static string get_global_param_string_buffer(string param)
    {
        EnsureCoolPropLoaded();
        IntPtr buffer = Marshal.AllocHGlobal(10000);
        try
        {
            long result = _get_global_param_string(param, buffer, 10000);
            if (result == 0)
                return null;
            return Marshal.PtrToStringAnsi(buffer);
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    private static string get_fluid_param_string(string fluid, string param)
    {
        EnsureCoolPropLoaded();
        IntPtr buffer = Marshal.AllocHGlobal(2000);
        try
        {
            long result = _get_fluid_param_string(fluid, param, buffer, 2000);
            if (result == 0)
                return GetCoolPropError();
            return Marshal.PtrToStringAnsi(buffer);
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
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
        IntPtr props1SIPtr = GetProcAddress(_coolPropHandle, "Props1SI");
        IntPtr phaseSIPtr = GetProcAddress(_coolPropHandle, "PhaseSI");
        IntPtr haPropsSIPtr = GetProcAddress(_coolPropHandle, "HAPropsSI");
        IntPtr getErrorPtr = GetProcAddress(_coolPropHandle, "get_global_param_string");
        IntPtr getFluidParamPtr = GetProcAddress(_coolPropHandle, "get_fluid_param_string");

        if (propsSIPtr != IntPtr.Zero)
            _PropsSI = Marshal.GetDelegateForFunctionPointer<PropsSI_Delegate>(propsSIPtr);
        if (props1SIPtr != IntPtr.Zero)
            _Props1SI = Marshal.GetDelegateForFunctionPointer<Props1SI_Delegate>(props1SIPtr);
        if (phaseSIPtr != IntPtr.Zero)
            _PhaseSI = Marshal.GetDelegateForFunctionPointer<PhaseSI_Delegate>(phaseSIPtr);
        if (haPropsSIPtr != IntPtr.Zero)
            _HAPropsSI = Marshal.GetDelegateForFunctionPointer<HAPropsSI_Delegate>(haPropsSIPtr);
        if (getErrorPtr != IntPtr.Zero)
            _get_global_param_string = Marshal.GetDelegateForFunctionPointer<get_global_param_string_Delegate>(getErrorPtr);
        if (getFluidParamPtr != IntPtr.Zero)
            _get_fluid_param_string = Marshal.GetDelegateForFunctionPointer<get_fluid_param_string_Delegate>(getFluidParamPtr);

        if (_PropsSI == null || _Props1SI == null || _HAPropsSI == null || _get_global_param_string == null)
        {
            throw new EntryPointNotFoundException("Required functions not found in CoolProp.dll");
        }
    }

    // Function to retrieve the error message from CoolProp
    private static string GetCoolPropError()
    {
        try
        {
            if (_get_global_param_string == null)
                return "CoolProp error function not initialized";
            
            string error = get_global_param_string_buffer("errstring");
            
            // Return the actual error message if available, otherwise a generic message
            if (string.IsNullOrEmpty(error))
                return "Unknown CoolProp error (no error details available)";
            
            return error.Trim();
        }
        catch (Exception ex)
        {
            return $"Error retrieving CoolProp error: {ex.Message}";
        }
    }

    // Function to calculate thermodynamic properties using SI units (no conversion)
    [ExcelFunction(Name = "PropsSI", Description = "Calculate thermodynamic properties of real fluids using CoolProp with SI units (K, Pa, J/kg, etc.).")]
    public static object PropsSI(string output, string name1, object value1, string name2, object value2, string fluid)
    {
        // Check for missing or invalid inputs
        if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
        if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
        if (value1 == null || !(value1 is double)) return "Error: First property value is missing or not a number.";
        if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
        if (value2 == null || !(value2 is double)) return "Error: Second property value is missing or not a number.";
        if (string.IsNullOrWhiteSpace(fluid)) return "Error: Fluid name is missing.";

        // Normalize parameter names and fluid name
        name1 = FormatName(name1);
        name2 = FormatName(name2);
        output = FormatName(output);
        fluid = FormatFluidName(fluid);

        // Call CoolProp directly with SI units (no conversion)
        double result;
        try
        {
            result = PropsSI_Internal(output, name1, (double)value1, name2, (double)value2, fluid);
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
            return $"Error: Exception occurred while computing property. {ex.Message}. CoolProp error: {GetCoolPropError()}";
        }

        // Check if the result indicates an error (CoolProp returns large values or NaN on error)
        if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
        {
            return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
        }

        // Return result in SI units (no conversion)
        return result;
    }

    // Function to calculate thermodynamic properties using engineering units
    [ExcelFunction(Name = "Props", Description = "Calculate thermodynamic properties of real fluids using CoolProp with engineering units (°C, bar, kJ/kg, etc.).")]
    public static object Props(string output, string name1, object value1, string name2, object value2, string fluid)
    {
        // Check for missing or invalid inputs
        if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
        if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
        if (value1 == null || !(value1 is double)) return "Error: First property value is missing or not a number.";
        if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
        if (value2 == null || !(value2 is double)) return "Error: Second property value is missing or not a number.";
        if (string.IsNullOrWhiteSpace(fluid)) return "Error: Fluid name is missing.";

        // Normalize parameter names and fluid name
        name1 = FormatName(name1);
        name2 = FormatName(name2);
        output = FormatName(output);
        fluid = FormatFluidName(fluid);

        // Convert inputs to SI units
        double val1SI = ConvertToSI(name1, (double)value1);
        double val2SI = ConvertToSI(name2, (double)value2);

        // Call CoolProp for the requested property
        double result;
        try
        {
            result = PropsSI_Internal(output, name1, val1SI, name2, val2SI, fluid);
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
            return $"Error: Exception occurred while computing property. {ex.Message}. CoolProp error: {GetCoolPropError()}";
        }

        // Check if the result indicates an error (CoolProp returns large values or NaN on error)
        if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
        {
            return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
        }

        // Convert output to engineering units
        return ConvertFromSI(output, result);
    }

    // TMPr alias for Props (uses engineering units)
    [ExcelFunction(Name = "TMPr", Description = "Calculate thermodynamic properties of real fluids using CoolProp with engineering units (°C, bar, kJ/kg, etc.). Alias for Props.")]
    public static object TMPr(string output, string name1, object value1, string name2, object value2, string fluid)
    {
        return Props(output, name1, value1, name2, value2, fluid);
    }

    [ExcelFunction(Name = "HAPropsSI", Description = "Calculate thermodynamic properties of humid air using CoolProp with SI units (K, Pa, J/kg, etc.).")]
    public static object HAPropsSI(string output, string name1, object value1, string name2, object value2, string name3, object value3)
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
            result = HAPropsSI_Internal(output, name1, (double)value1, name2, (double)value2, name3, (double)value3);
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
            return $"Error: Exception occurred while computing property. {ex.Message}. CoolProp error: {GetCoolPropError()}";
        }

        // Check if the result indicates an error (CoolProp returns large values or NaN on error)
        if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
        {
            return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
        }

        // Return result in SI units (no conversion)
        return result;
    }

    [ExcelFunction(Name = "HAProps", Description = "Calculate thermodynamic properties of humid air using CoolProp with engineering units (°C, bar, kJ/kg, etc.).")]
    public static object HAProps(string output, string name1, object value1, string name2, object value2, string name3, object value3)
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
            result = HAPropsSI_Internal(output, name1, val1SI, name2, val2SI, name3, val3SI);
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
            return $"Error: Exception occurred while computing property. {ex.Message}. CoolProp error: {GetCoolPropError()}";
        }

        // Check if the result indicates an error (CoolProp returns large values or NaN on error)
        if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
        {
            return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
        }

        // Convert output to engineering units
        return ConvertFromSI_HA(output, result);
    }

    // TMPa alias for HAProps (uses engineering units)
    [ExcelFunction(Name = "TMPa", Description = "Calculate thermodynamic properties of humid air using CoolProp with engineering units (°C, bar, kJ/kg, etc.). Alias for HAProps.")]
    public static object TMPa(string output, string name1, object value1, string name2, object value2, string name3, object value3)
    {
        return HAProps(output, name1, value1, name2, value2, name3, value3);
    }

    // Function to get phase using SI units
    [ExcelFunction(Name = "PhaseSI", Description = "Get the phase of a fluid using CoolProp with SI units (K, Pa).")]
    public static object PhaseSI(string name1, object value1, string name2, object value2, string fluid)
    {
        // Check for missing or invalid inputs
        if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
        if (value1 == null || !(value1 is double)) return "Error: First property value is missing or not a number.";
        if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
        if (value2 == null || !(value2 is double)) return "Error: Second property value is missing or not a number.";
        if (string.IsNullOrWhiteSpace(fluid)) return "Error: Fluid name is missing.";

        // Normalize parameter names and fluid name
        name1 = FormatName(name1);
        name2 = FormatName(name2);
        fluid = FormatFluidName(fluid);

        try
        {
            string phase = PhaseSI_Internal(name1, (double)value1, name2, (double)value2, fluid);
            if (string.IsNullOrEmpty(phase))
                return $"Error: CoolProp failed to determine phase. {GetCoolPropError()}";
            return phase;
        }
        catch (Exception ex)
        {
            return $"Error: Exception occurred while determining phase. {ex.Message}. CoolProp error: {GetCoolPropError()}";
        }
    }

    // Function to get phase using engineering units
    [ExcelFunction(Name = "Phase", Description = "Get the phase of a fluid using CoolProp with engineering units (°C, bar).")]
    public static object Phase(string name1, object value1, string name2, object value2, string fluid)
    {
        // Check for missing or invalid inputs
        if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
        if (value1 == null || !(value1 is double)) return "Error: First property value is missing or not a number.";
        if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
        if (value2 == null || !(value2 is double)) return "Error: Second property value is missing or not a number.";
        if (string.IsNullOrWhiteSpace(fluid)) return "Error: Fluid name is missing.";

        // Normalize parameter names and fluid name
        name1 = FormatName(name1);
        name2 = FormatName(name2);
        fluid = FormatFluidName(fluid);

        // Convert inputs to SI units
        double val1SI = ConvertToSI(name1, (double)value1);
        double val2SI = ConvertToSI(name2, (double)value2);

        try
        {
            string phase = PhaseSI_Internal(name1, val1SI, name2, val2SI, fluid);
            if (string.IsNullOrEmpty(phase))
                return $"Error: CoolProp failed to determine phase. {GetCoolPropError()}";
            return phase;
        }
        catch (Exception ex)
        {
            return $"Error: Exception occurred while determining phase. {ex.Message}. CoolProp error: {GetCoolPropError()}";
        }
    }

    // Function to calculate single-input properties using SI units
    [ExcelFunction(Name = "Props1SI", Description = "Calculate single-input properties (critical properties, molar mass, etc.) using CoolProp with SI units (K, Pa, kg/mol, etc.).")]
    public static object Props1SI(string output, string fluid)
    {
        // Check for missing or invalid inputs
        if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
        if (string.IsNullOrWhiteSpace(fluid)) return "Error: Fluid name is missing.";

        // Normalize parameter name and fluid name
        output = FormatName(output);
        fluid = FormatFluidName(fluid);

        // Call CoolProp directly with SI units (no conversion)
        double result;
        try
        {
            result = Props1SI_Internal(output, fluid);
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
            return $"Error: Exception occurred while computing property. {ex.Message}. CoolProp error: {GetCoolPropError()}";
        }

        // Check if the result indicates an error (CoolProp returns large values or NaN on error)
        if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
        {
            return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
        }

        // Return result in SI units (no conversion)
        return result;
    }

    // Function to calculate single-input properties using engineering units
    [ExcelFunction(Name = "Props1", Description = "Calculate single-input properties (critical properties, molar mass, etc.) using CoolProp with engineering units (°C, bar, kg/mol, etc.).")]
    public static object Props1(string output, string fluid)
    {
        // Check for missing or invalid inputs
        if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
        if (string.IsNullOrWhiteSpace(fluid)) return "Error: Fluid name is missing.";

        // Normalize parameter name and fluid name
        output = FormatName(output);
        fluid = FormatFluidName(fluid);

        // Call CoolProp for the requested property
        double result;
        try
        {
            result = Props1SI_Internal(output, fluid);
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
            return $"Error: Exception occurred while computing property. {ex.Message}. CoolProp error: {GetCoolPropError()}";
        }

        // Check if the result indicates an error (CoolProp returns large values or NaN on error)
        if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
        {
            return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
        }

        // Convert output to engineering units
        return ConvertFromSI(output, result);
    }

    // Function to get CoolProp version and global parameters
    [ExcelFunction(Name = "GetGlobalParam", Description = "Get CoolProp version, revision, or other global parameter strings.")]
    public static object GetGlobalParam(string param)
    {
        // Check for missing input
        if (string.IsNullOrWhiteSpace(param)) return "Error: Parameter name is missing.";

        try
        {
            string result = get_global_param_string_buffer(param);
            if (string.IsNullOrEmpty(result))
                return $"Error: Failed to retrieve global parameter '{param}'. Parameter may not exist or be invalid.";
            
            return result.Trim();
        }
        catch (Exception ex)
        {
            return $"Error: Exception occurred while retrieving global parameter. {ex.Message}";
        }
    }

    // Function to get fluid-specific parameter strings
    [ExcelFunction(Name = "GetFluidParam", Description = "Get fluid-specific parameter strings (aliases, CAS number, REFPROP name, etc.).")]
    public static object GetFluidParam(string fluid, string param)
    {
        // Check for missing inputs
        if (string.IsNullOrWhiteSpace(fluid)) return "Error: Fluid name is missing.";
        if (string.IsNullOrWhiteSpace(param)) return "Error: Parameter name is missing.";

        // Normalize fluid name
        fluid = FormatFluidName(fluid);

        try
        {
            string result = get_fluid_param_string(fluid, param);
            if (string.IsNullOrEmpty(result))
                return $"Error: Failed to retrieve fluid parameter '{param}' for fluid '{fluid}'. {GetCoolPropError()}";
            
            return result;
        }
        catch (Exception ex)
        {
            return $"Error: Exception occurred while retrieving fluid parameter. {ex.Message}";
        }
    }

    // Function to create mixture string from element and composition ranges
    [ExcelFunction(Name = "MixtureString", Description = "Create a mixture string for CoolProp from element names and their mole/mass fractions.")]
    public static object MixtureString(object[,] elements, object[,] fractions)
    {
        try
        {
            // Validate inputs
            if (elements == null || fractions == null)
                return "Error: Elements and fractions arrays cannot be null.";

            int elemCount = elements.GetLength(0);
            int fracCount = fractions.GetLength(0);

            if (elemCount != fracCount)
                return $"Error: Number of elements ({elemCount}) must match number of fractions ({fracCount}).";

            if (elemCount == 0)
                return "Error: At least one element is required.";

            // Build mixture string
            var components = new List<string>();
            for (int i = 0; i < elemCount; i++)
            {
                // Get element name (handle single column or row)
                object elemValue = elements.GetLength(1) == 1 ? elements[i, 0] : elements[0, i];
                if (elemValue == null || string.IsNullOrWhiteSpace(elemValue.ToString()))
                    continue;

                string element = elemValue.ToString().Trim();

                // Get fraction value (handle single column or row)
                object fracValue = fractions.GetLength(1) == 1 ? fractions[i, 0] : fractions[0, i];
                if (fracValue == null || !(fracValue is double))
                    return $"Error: Fraction value at position {i + 1} is not a number.";

                double fraction = (double)fracValue;
                if (fraction < 0 || fraction > 1)
                    return $"Error: Fraction value {fraction} at position {i + 1} is out of range [0, 1].";

                components.Add($"{element}[{fraction}]");
            }

            if (components.Count == 0)
                return "Error: No valid components found.";

            // Return mixture string in CoolProp format: HEOS::Fluid1[fraction1]&Fluid2[fraction2]
            return "HEOS::" + string.Join("&", components);
        }
        catch (Exception ex)
        {
            return $"Error creating mixture string: {ex.Message}";
        }
    }


    // Normalize name capitalization to match the expected format (comprehensive parameter list)
    private static string FormatName(string name)
    {
        switch (name.ToUpper())
        {
            // Temperature
            case "T":
            case "TEMP":
            case "TEMPERATURE":
                return "T";
            
            // Pressure
            case "P":
            case "PRES":
            case "PRESSURE":
                return "P";
            
            // Enthalpy
            case "H":
            case "ENTH":
            case "ENTHALPY":
            case "HMASS":
            case "H_MASS":
                return "H";
            case "HMOLAR":
            case "H_MOLAR":
                return "Hmolar";
            
            // Internal Energy
            case "U":
            case "INTERNALENERGY":
            case "UMASS":
            case "U_MASS":
                return "U";
            case "UMOLAR":
            case "U_MOLAR":
                return "Umolar";
            
            // Entropy
            case "S":
            case "ENTR":
            case "ENTROPY":
            case "SMASS":
            case "S_MASS":
                return "S";
            case "SMOLAR":
            case "S_MOLAR":
                return "Smolar";
            case "SMOLAR_RESIDUAL":
            case "SMOLARRESIDUAL":
                return "Smolar_residual";
            
            // Density
            case "D":
            case "RHO":
            case "DENS":
            case "DENSITY":
            case "DMASS":
            case "D_MASS":
            case "RHOMASS":
                return "D";
            case "DMOLAR":
            case "RHOMOLAR":
            case "DMOL":
            case "D_MOL":
            case "D_MOLAR":
                return "Dmolar";
            
            // Specific Heat Capacity
            case "CVMASS":
            case "CV":
                return "Cvmass";
            case "CPMASS":
            case "CP":
            case "C":
                return "Cpmass";
            case "CPMOLAR":
            case "CPMOL":
                return "Cpmolar";
            case "CVMOLAR":
            case "CVMOL":
                return "Cvmolar";
            case "CP0MASS":
                return "Cp0mass";
            case "CP0MOLAR":
                return "Cp0molar";
            
            // Quality (vapor fraction)
            case "Q":
            case "QUALITY":
            case "X":
            case "VAPORFRACTION":
            case "VAPOR_FRACTION":
                return "Q";
            
            // Reduced properties
            case "TAU":
                return "Tau";
            case "DELTA":
                return "Delta";
            
            // Helmholtz energy derivatives
            case "ALPHA0":
                return "Alpha0";
            case "ALPHAR":
                return "Alphar";
            case "DALPHA0_DDELTA_CONSTTAU":
            case "DALPHA0_DDELTA":
                return "DALPHA0_DDELTA_CONSTTAU";
            case "DALPHA0_DTAU_CONSTDELTA":
            case "DALPHA0_DTAU":
                return "DALPHA0_DTAU_CONSTDELTA";
            case "DALPHAR_DDELTA_CONSTTAU":
            case "DALPHAR_DDELTA":
                return "DALPHAR_DDELTA_CONSTTAU";
            case "DALPHAR_DTAU_CONSTDELTA":
            case "DALPHAR_DTAU":
                return "DALPHAR_DTAU_CONSTDELTA";
            
            // Speed of sound
            case "SPEED_OF_SOUND":
            case "SPEEDOFSOUND":
            case "A":
            case "W":
                return "A";
            
            // Virial coefficients
            case "BVIRIAL":
                return "Bvirial";
            case "CVIRIAL":
                return "Cvirial";
            case "DBVIRIAL_DT":
            case "DBVIRIAL":
            case "DBVIRIALDT":
                return "DBVIRIAL_DT";
            case "DCVIRIAL_DT":
            case "DCVIRIAL":
            case "DCVIRIALDT":
                return "DCVIRIAL_DT";
            
            // Conductivity
            case "K":
            case "CONDUCTIVITY":
            case "L":
                return "conductivity";
            
            // Dipole moment
            case "DIPOLE_MOMENT":
            case "DIPOLEMOMENT":
            case "DIPOLE":
                return "DIPOLE_MOMENT";
            
            // Gibbs energy
            case "G":
            case "GMASS":
            case "GIBBS":
                return "G";
            case "GMOLAR":
                return "Gmolar";
            
            // Helmholtz energy
            case "HELMHOLTZMASS":
                return "HELMHOLTZMASS";
            case "HELMHOLTZMOLAR":
                return "HELMHOLTZMOLAR";
            
            // Isentropic expansion coefficient
            case "GAMMA":
            case "ISENTROPICEXPANSIONCOEFFICIENT":
                return "isentropic_expansion_coefficient";
            
            // Expansion and compressibility
            case "ISOBARIC_EXPANSION_COEFFICIENT":
            case "ISOBARICEXPANSION":
            case "ISOBARICEXPANSIONCOEFFICIENT":
                return "isobaric_expansion_coefficient";
            case "ISOTHERMAL_COMPRESSIBILITY":
            case "ISOTHERMALCOMPRESSIBILITY":
                return "isothermal_compressibility";
            
            // Surface tension
            case "SURFACE_TENSION":
            case "SURFACETENSION":
            case "SIGMA":
            case "I":
                return "surface_tension";
            
            // Molar mass
            case "MM":
            case "MOLAR_MASS":
            case "MOLARMASS":
            case "MOLEMASS":
                return "M";
            
            // Critical properties
            case "PCRIT":
            case "P_CRITICAL":
            case "PCRITICAL":
                return "Pcrit";
            case "TCRIT":
            case "T_CRITICAL":
            case "TCRITICAL":
                return "Tcrit";
            case "RHOCRIT":
            case "RHOCRITICAL":
            case "DCRIT":
                return "rhocrit";
            case "RHOMASS_CRITICAL":
            case "RHOMASSCRITICAL":
                return "rhomass_critical";
            case "RHOMOLAR_CRITICAL":
            case "RHOMOLARCRITICAL":
                return "rhomolar_critical";
            
            // Phase
            case "PHASE":
                return "Phase";
            
            // Max/Min properties
            case "PMAX":
            case "P_MAX":
                return "pmax";
            case "PMIN":
            case "P_MIN":
                return "pmin";
            case "TMAX":
            case "T_MAX":
                return "Tmax";
            case "TMIN":
            case "T_MIN":
                return "Tmin";
            
            // Prandtl number
            case "PRANDTL":
                return "Prandtl";
            
            // Triple point
            case "PTRIPLE":
            case "P_TRIPLE":
            case "PTRIP":
                return "ptriple";
            case "TTRIPLE":
            case "T_TRIPLE":
            case "TTRIP":
                return "Ttriple";
            
            // Reducing properties
            case "P_REDUCING":
            case "PREDUCING":
                return "p_reducing";
            case "T_REDUCING":
            case "TREDUCING":
                return "T_reducing";
            case "RHOMASS_REDUCING":
            case "RHOMASSREDUCING":
                return "rhomass_reducing";
            case "RHOMOLAR_REDUCING":
            case "RHOMOLARREDUCING":
                return "rhomolar_reducing";
            
            // Freeze temperature
            case "T_FREEZE":
            case "TFREEZE":
            case "FREEZING_TEMPERATURE":
                return "T_freeze";
            
            // Viscosity
            case "MU":
            case "VISCOSITY":
            case "V":
                return "viscosity";
            
            // Compressibility factor
            case "Z":
            case "COMPRESSIBILITY":
            case "COMPRESSIBILITYFACTOR":
                return "Z";
            
            // Acentric factor
            case "ACENTRIC":
            case "ACENTRIC_FACTOR":
            case "ACENTRICFACTOR":
            case "OMEGA":
                return "acentric";
            
            // Fundamental derivative of gas dynamics
            case "FUNDAMENTAL_DERIVATIVE_OF_GAS_DYNAMICS":
            case "FUNDAMENTALDERIVATIVE":
            case "FH":
                return "fundamental_derivative_of_gas_dynamics";
            
            // Gas constant
            case "GAS_CONSTANT":
            case "GASCONSTANT":
                return "gas_constant";
            
            // Fraction limits
            case "FRACTION_MAX":
            case "FRACTIONMAX":
                return "fraction_max";
            case "FRACTION_MIN":
            case "FRACTIONMIN":
                return "fraction_min";
            
            // Global warming potential
            case "GWP20":
            case "GWP_20":
                return "GWP20";
            case "GWP100":
            case "GWP_100":
                return "GWP100";
            case "GWP500":
            case "GWP_500":
                return "GWP500";
            
            // Ozone depletion potential
            case "ODP":
            case "OZONEDEPLETIONPOTENTIAL":
                return "ODP";
            
            // Helmholtz energy (aliases)
            case "HH":
                return "HELMHOLTZMASS";
            
            // Phase identification parameter
            case "PIP":
                return "PIP";
            
            default:
                return name; // Return as is if no match is found
        }
    }

    // Normalize fluid names to match CoolProp's expected format (case-insensitive)
    private static string FormatFluidName(string fluid)
    {
        // Handle mixture strings (contain "&" or "HEOS::")
        if (fluid.Contains("&") || fluid.ToUpper().StartsWith("HEOS::"))
            return fluid;

        string upper = fluid.ToUpper().Replace("-", "").Replace("_", "").Replace(" ", "");
        
        // Pure fluids (alphabetical)
        switch (upper)
        {
            case "1BUTENE": return "1-Butene";
            case "ACETONE": return "Acetone";
            case "AIR": return "Air";
            case "AMMONIA": case "NH3": return "Ammonia";
            case "ARGON": case "AR": return "Argon";
            case "BENZENE": return "Benzene";
            case "CARBONDIOXIDE": case "CO2": return "CarbonDioxide";
            case "CARBONMONOXIDE": case "CO": return "CarbonMonoxide";
            case "CARBONYLSULFIDE": case "COS": return "CarbonylSulfide";
            case "CIS2BUTENE": return "cis-2-Butene";
            case "CYCLOHEXANE": return "CycloHexane";
            case "CYCLOPENTANE": return "Cyclopentane";
            case "CYCLOPROPANE": return "CycloPropane";
            case "D4": return "D4";
            case "D5": return "D5";
            case "D6": return "D6";
            case "DEUTERIUM": case "D2": return "Deuterium";
            case "DICHLOROETHANE": return "Dichloroethane";
            case "DIETHYLETHER": return "DiethylEther";
            case "DIMETHYLCARBONATE": return "DimethylCarbonate";
            case "DIMETHYLETHER": case "DME": return "DimethylEther";
            case "ETHANE": case "C2H6": return "Ethane";
            case "ETHANOL": return "Ethanol";
            case "ETHYLBENZENE": return "EthylBenzene";
            case "ETHYLENE": case "C2H4": return "Ethylene";
            case "ETHYLENEOXIDE": return "EthyleneOxide";
            case "FLUORINE": case "F2": return "Fluorine";
            case "HEAVYWATER": case "D2O": return "HeavyWater";
            case "HELIUM": case "HE": return "Helium";
            case "HFE143M": return "HFE143m";
            case "HYDROGEN": case "H2": return "Hydrogen";
            case "HYDROGENCHLORIDE": case "HCL": return "HydrogenChloride";
            case "HYDROGENSULFIDE": case "H2S": return "HydrogenSulfide";
            case "ISOBUTANE": case "IBUTANE": return "IsoButane";
            case "ISOBUTENE": case "IBUTENE": return "IsoButene";
            case "ISOHEXANE": return "Isohexane";
            case "ISOPENTANE": return "Isopentane";
            case "KRYPTON": case "KR": return "Krypton";
            case "MXYLENE": return "m-Xylene";
            case "MD2M": return "MD2M";
            case "MD3M": return "MD3M";
            case "MD4M": return "MD4M";
            case "MDM": return "MDM";
            case "METHANE": case "CH4": return "Methane";
            case "METHANOL": case "MEOH": return "Methanol";
            case "METHYLLINOLEATE": return "MethylLinoleate";
            case "METHYLLINOLENATE": return "MethylLinolenate";
            case "METHYLOLEATE": return "MethylOleate";
            case "METHYLPALMITATE": return "MethylPalmitate";
            case "METHYLSTEARATE": return "MethylStearate";
            case "MM": return "MM";
            case "NBUTANE": case "BUTANE": return "n-Butane";
            case "NDECANE": case "DECANE": return "n-Decane";
            case "NDODECANE": case "DODECANE": return "n-Dodecane";
            case "NHEPTANE": case "HEPTANE": return "n-Heptane";
            case "NHEXANE": case "HEXANE": return "n-Hexane";
            case "NNONANE": case "NONANE": return "n-Nonane";
            case "NOCTANE": case "OCTANE": return "n-Octane";
            case "NPENTANE": case "PENTANE": return "n-Pentane";
            case "NPROPANE": case "PROPANE": return "n-Propane";
            case "NUNDECANE": case "UNDECANE": return "n-Undecane";
            case "NEON": case "NE": return "Neon";
            case "NEOPENTANE": return "Neopentane";
            case "NITROGEN": case "N2": return "Nitrogen";
            case "NITROUSOXIDE": case "N2O": return "NitrousOxide";
            case "NOVEC649": return "Novec649";
            case "OXYLENE": return "o-Xylene";
            case "ORTHODEUTERIUM": return "OrthoDeuterium";
            case "ORTHOHYDROGEN": return "OrthoHydrogen";
            case "OXYGEN": case "O2": return "Oxygen";
            case "PXYLENE": return "p-Xylene";
            case "PARADEUTERIUM": return "ParaDeuterium";
            case "PARAHYDROGEN": return "ParaHydrogen";
            case "PROPYLENE": case "C3H6": return "Propylene";
            case "PROPYNE": return "Propyne";
            case "R11": return "R11";
            case "R113": return "R113";
            case "R114": return "R114";
            case "R115": return "R115";
            case "R116": return "R116";
            case "R12": return "R12";
            case "R123": return "R123";
            case "R1233ZDE": case "R1233ZD(E)": return "R1233zd(E)";
            case "R1234YF": return "R1234yf";
            case "R1234ZEE": case "R1234ZE(E)": return "R1234ze(E)";
            case "R1234ZEZ": case "R1234ZE(Z)": return "R1234ze(Z)";
            case "R124": return "R124";
            case "R125": return "R125";
            case "R13": return "R13";
            case "R134A": return "R134a";
            case "R13I1": return "R13I1";
            case "R14": return "R14";
            case "R141B": return "R141b";
            case "R142B": return "R142b";
            case "R143A": return "R143a";
            case "R152A": return "R152A";
            case "R161": return "R161";
            case "R21": return "R21";
            case "R218": return "R218";
            case "R22": return "R22";
            case "R227EA": return "R227EA";
            case "R23": return "R23";
            case "R236EA": return "R236EA";
            case "R236FA": return "R236FA";
            case "R245CA": return "R245ca";
            case "R245FA": return "R245fa";
            case "R32": return "R32";
            case "R365MFC": return "R365MFC";
            case "R40": return "R40";
            case "R404A": return "R404A";
            case "R407C": return "R407C";
            case "R41": return "R41";
            case "R410A": return "R410A";
            case "R507A": return "R507A";
            case "RC318": return "RC318";
            case "SES36": return "SES36";
            case "SULFURDIOXIDE": case "SO2": return "SulfurDioxide";
            case "SULFURHEXAFLUORIDE": case "SF6": return "SulfurHexafluoride";
            case "TOLUENE": return "Toluene";
            case "TRANS2BUTENE": return "trans-2-Butene";
            case "WATER": case "H2O": return "Water";
            case "XENON": case "XE": return "Xenon";
            
            // Predefined mixtures
            case "AIR.MIX": return "Air.mix";
            case "AMARILLO.MIX": return "Amarillo.mix";
            case "EKOFISK.MIX": return "Ekofisk.mix";
            case "GULFCOAST.MIX": return "GulfCoast.mix";
            case "GULFCOASTGAS(NIST1).MIX": case "GULFCOASTGASNIST1.MIX": return "GulfCoastGas(NIST1).mix";
            case "HIGHCO2.MIX": return "HighCO2.mix";
            case "HIGHN2.MIX": return "HighN2.mix";
            case "NATURALGASSAMPLE.MIX": return "NaturalGasSample.mix";
            case "R401A.MIX": return "R401A.mix";
            case "R401B.MIX": return "R401B.mix";
            case "R401C.MIX": return "R401C.mix";
            case "R402A.MIX": return "R402A.mix";
            case "R402B.MIX": return "R402B.mix";
            case "R403A.MIX": return "R403A.mix";
            case "R403B.MIX": return "R403B.mix";
            case "R404A.MIX": return "R404A.mix";
            case "R405A.MIX": return "R405A.mix";
            case "R406A.MIX": return "R406A.mix";
            case "R407A.MIX": return "R407A.mix";
            case "R407B.MIX": return "R407B.mix";
            case "R407C.MIX": return "R407C.mix";
            case "R407D.MIX": return "R407D.mix";
            case "R407E.MIX": return "R407E.mix";
            case "R407F.MIX": return "R407F.mix";
            case "R408A.MIX": return "R408A.mix";
            case "R409A.MIX": return "R409A.mix";
            case "R409B.MIX": return "R409B.mix";
            case "R410A.MIX": return "R410A.mix";
            case "R410B.MIX": return "R410B.mix";
            case "R411A.MIX": return "R411A.mix";
            case "R411B.MIX": return "R411B.mix";
            case "R412A.MIX": return "R412A.mix";
            case "R413A.MIX": return "R413A.mix";
            case "R414A.MIX": return "R414A.mix";
            case "R414B.MIX": return "R414B.mix";
            case "R415A.MIX": return "R415A.mix";
            case "R415B.MIX": return "R415B.mix";
            case "R416A.MIX": return "R416A.mix";
            case "R417A.MIX": return "R417A.mix";
            case "R417B.MIX": return "R417B.mix";
            case "R417C.MIX": return "R417C.mix";
            case "R418A.MIX": return "R418A.mix";
            case "R419A.MIX": return "R419A.mix";
            case "R419B.MIX": return "R419B.mix";
            case "R420A.MIX": return "R420A.mix";
            case "R421A.MIX": return "R421A.mix";
            case "R421B.MIX": return "R421B.mix";
            case "R422A.MIX": return "R422A.mix";
            case "R422B.MIX": return "R422B.mix";
            case "R422C.MIX": return "R422C.mix";
            case "R422D.MIX": return "R422D.mix";
            case "R422E.MIX": return "R422E.mix";
            case "R423A.MIX": return "R423A.mix";
            case "R424A.MIX": return "R424A.mix";
            case "R425A.MIX": return "R425A.mix";
            case "R426A.MIX": return "R426A.mix";
            case "R427A.MIX": return "R427A.mix";
            case "R428A.MIX": return "R428A.mix";
            case "R429A.MIX": return "R429A.mix";
            case "R430A.MIX": return "R430A.mix";
            case "R431A.MIX": return "R431A.mix";
            case "R432A.MIX": return "R432A.mix";
            case "R433A.MIX": return "R433A.mix";
            case "R433B.MIX": return "R433B.mix";
            case "R433C.MIX": return "R433C.mix";
            case "R434A.MIX": return "R434A.mix";
            case "R435A.MIX": return "R435A.mix";
            case "R436A.MIX": return "R436A.mix";
            case "R436B.MIX": return "R436B.mix";
            case "R437A.MIX": return "R437A.mix";
            case "R438A.MIX": return "R438A.mix";
            case "R439A.MIX": return "R439A.mix";
            case "R440A.MIX": return "R440A.mix";
            case "R441A.MIX": return "R441A.mix";
            case "R442A.MIX": return "R442A.mix";
            case "R443A.MIX": return "R443A.mix";
            case "R444A.MIX": return "R444A.mix";
            case "R444B.MIX": return "R444B.mix";
            case "R445A.MIX": return "R445A.mix";
            case "R446A.MIX": return "R446A.mix";
            case "R447A.MIX": return "R447A.mix";
            case "R448A.MIX": return "R448A.mix";
            case "R449A.MIX": return "R449A.mix";
            case "R449B.MIX": return "R449B.mix";
            case "R450A.MIX": return "R450A.mix";
            case "R451A.MIX": return "R451A.mix";
            case "R451B.MIX": return "R451B.mix";
            case "R452A.MIX": return "R452A.mix";
            case "R453A.MIX": return "R453A.mix";
            case "R454A.MIX": return "R454A.mix";
            case "R454B.MIX": return "R454B.mix";
            case "R500.MIX": return "R500.mix";
            case "R501.MIX": return "R501.mix";
            case "R502.MIX": return "R502.mix";
            case "R503.MIX": return "R503.mix";
            case "R504.MIX": return "R504.mix";
            case "R507A.MIX": return "R507A.mix";
            case "R508A.MIX": return "R508A.mix";
            case "R508B.MIX": return "R508B.mix";
            case "R509A.MIX": return "R509A.mix";
            case "R510A.MIX": return "R510A.mix";
            case "R511A.MIX": return "R511A.mix";
            case "R512A.MIX": return "R512A.mix";
            case "R513A.MIX": return "R513A.mix";
            case "TYPICALNATURALGAS.MIX": return "TypicalNaturalGas.mix";
            
            default:
                return fluid; // Return as is if no match is found
        }
    }

    // Convert from custom units to SI units
    private static double ConvertToSI(string name, double value)
    {
        switch (name)
        {
            case "T": // Temperature (°C to K)
            case "Tcrit": // Critical temperature (°C to K)
            case "Tmax": // Maximum temperature (°C to K)
            case "Tmin": // Minimum temperature (°C to K)
            case "Ttriple": // Triple point temperature (°C to K)
            case "T_freeze": // Freeze temperature (°C to K)
            case "T_reducing": // Reducing temperature (°C to K)
                return value + 273.15;
            case "P": // Pressure (bar to Pa)
            case "Pcrit": // Critical pressure (bar to Pa)
            case "pmax": // Maximum pressure (bar to Pa)
            case "pmin": // Minimum pressure (bar to Pa)
            case "ptriple": // Triple point pressure (bar to Pa)
            case "p_reducing": // Reducing pressure (bar to Pa)
                return value * 1e5;
            case "H": // Specific Enthalpy (kJ/kg to J/kg)
            case "Hmolar": // Molar Enthalpy (kJ/mol to J/mol)
            case "U": // Specific Internal Energy (kJ/kg to J/kg)
            case "Umolar": // Molar Internal Energy (kJ/mol to J/mol)
            case "S": // Specific Entropy (kJ/kg/K to J/kg/K)
            case "Smolar": // Molar Entropy (kJ/mol/K to J/mol/K)
            case "Smolar_residual": // Residual molar entropy (kJ/mol/K to J/mol/K)
            case "Cpmass": // Mass specific heat at constant pressure (kJ/kg/K to J/kg/K)
            case "Cvmass": // Mass specific heat at constant volume (kJ/kg/K to J/kg/K)
            case "Cpmolar": // Molar specific heat at constant pressure (kJ/mol/K to J/mol/K)
            case "Cvmolar": // Molar specific heat at constant volume (kJ/mol/K to J/mol/K)
            case "Cp0mass": // Ideal gas mass specific heat (kJ/kg/K to J/kg/K)
            case "Cp0molar": // Ideal gas molar specific heat (kJ/mol/K to J/mol/K)
            case "G": // Gibbs energy (kJ/kg to J/kg)
            case "Gmolar": // Molar Gibbs energy (kJ/mol to J/mol)
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
            case "Tcrit": // Critical temperature (K to °C)
            case "Tmax": // Maximum temperature (K to °C)
            case "Tmin": // Minimum temperature (K to °C)
            case "Ttriple": // Triple point temperature (K to °C)
            case "T_freeze": // Freeze temperature (K to °C)
            case "T_reducing": // Reducing temperature (K to °C)
                return value - 273.15;
            case "P": // Pressure (Pa to bar)
            case "Pcrit": // Critical pressure (Pa to bar)
            case "pmax": // Maximum pressure (Pa to bar)
            case "pmin": // Minimum pressure (Pa to bar)
            case "ptriple": // Triple point pressure (Pa to bar)
            case "p_reducing": // Reducing pressure (Pa to bar)
                return value / 1e5;
            case "H": // Specific Enthalpy (J/kg to kJ/kg)
            case "Hmolar": // Molar Enthalpy (J/mol to kJ/mol)
            case "U": // Specific Internal Energy (J/kg to kJ/kg)
            case "Umolar": // Molar Internal Energy (J/mol to kJ/mol)
            case "S": // Specific Entropy (J/kg/K to kJ/kg/K)
            case "Smolar": // Molar Entropy (J/mol/K to kJ/mol/K)
            case "Smolar_residual": // Residual molar entropy (J/mol/K to kJ/mol/K)
            case "Cpmass": // Mass specific heat at constant pressure (J/kg/K to kJ/kg/K)
            case "Cvmass": // Mass specific heat at constant volume (J/kg/K to kJ/kg/K)
            case "Cpmolar": // Molar specific heat at constant pressure (J/mol/K to kJ/mol/K)
            case "Cvmolar": // Molar specific heat at constant volume (J/mol/K to kJ/mol/K)
            case "Cp0mass": // Ideal gas mass specific heat (J/kg/K to kJ/kg/K)
            case "Cp0molar": // Ideal gas molar specific heat (J/mol/K to kJ/mol/K)
            case "G": // Gibbs energy (J/kg to kJ/kg)
            case "Gmolar": // Molar Gibbs energy (J/mol to kJ/mol)
                return value / 1000;
            default:
                return value; // No conversion if unit not found
        }
    }

    // Normalize name capitalization for HA properties (comprehensive with case-insensitive support)
    private static string FormatName_HA(string name)
    {
        switch (name.ToUpper())
        {
            // Wet Bulb Temperature
            case "B":
            case "TWB":
            case "T_WB":
            case "WETBULB":
            case "WETBULBTEMP":
            case "WETBULBTEMPERATURE":
                return "Twb";
            
            // Specific heat per unit dry air (cp)
            case "C":
            case "CP":
            case "CDA":
            case "CPDA":
            case "CP_DA":
                return "Cda";
            
            // Specific heat of humid air
            case "CHA":
            case "CPHA":
            case "CP_HA":
                return "Cha";
            
            // Constant volume specific heat per unit dry air
            case "CV":
            case "CVMASS":
                return "CV";
            
            // Constant volume specific heat per unit humid air
            case "CVHA":
            case "CV_HA":
                return "CVha";
            
            // Dew Point Temperature
            case "D":
            case "TDP":
            case "T_DP":
            case "DEWPOINT":
            case "DEWPOINTTEMP":
            case "DEWPOINTTEMPERATURE":
                return "Tdp";
            
            // Enthalpy of dry air
            case "H":
            case "HDA":
            case "H_DA":
            case "ENTHALPY":
                return "Hda";
            
            // Enthalpy of humid air
            case "HHA":
            case "H_HA":
                return "Hha";
            
            // Thermal conductivity
            case "K":
            case "CONDUCTIVITY":
            case "THERMALCONDUCTIVITY":
                return "K";
            
            // Dynamic viscosity
            case "M":
            case "MU":
            case "VISC":
            case "VISCOSITY":
            case "DYNAMICVISCOSITY":
                return "MU";
            
            // Water mole fraction
            case "PSI_W":
            case "PSIW":
            case "Y":
                return "Psi_w";
            
            // Pressure
            case "P":
            case "PRESSURE":
            case "PRES":
                return "P";
            
            // Partial pressure of water
            case "P_W":
            case "PW":
            case "PARTIALPRESSURE":
            case "WATERPRESSURE":
                return "P_w";
            
            // Relative humidity
            case "R":
            case "RH":
            case "RELHUM":
            case "RELATIVEHUMIDITY":
                return "R";
            
            // Entropy of dry air
            case "S":
            case "SDA":
            case "S_DA":
            case "ENTROPY":
                return "Sda";
            
            // Entropy of humid air
            case "SHA":
            case "S_HA":
                return "Sha";
            
            // Dry bulb temperature
            case "T":
            case "TDB":
            case "T_DB":
            case "TEMP":
            case "TEMPERATURE":
            case "DRYBULB":
            case "DRYBULBTEMP":
            case "DRYBULBTEMPERATURE":
                return "T";
            
            // Specific volume of dry air
            case "V":
            case "VDA":
            case "V_DA":
                return "Vda";
            
            // Specific volume of humid air
            case "VHA":
            case "V_HA":
                return "Vha";
            
            // Humidity ratio
            case "W":
            case "OMEGA":
            case "HUMRAT":
            case "HUMIDITYRATIO":
            case "MIXINGRATIO":
                return "W";
            
            // Compressibility factor
            case "Z":
            case "COMPRESSIBILITY":
            case "COMPRESSIBILITYFACTOR":
                return "Z";
            
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
            case "H": // Specific Enthalpy of humid air (kJ/kg to J/kg)
            case "Sda": // Specific entropy of dry air (kJ/kg/K to J/kg/K)
            case "Sha": // Specific entropy of humid air (kJ/kg/K to J/kg/K)
            case "S": // Specific entropy of humid air (kJ/kg/K to J/kg/K)
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
            case "H": // Specific Enthalpy of humid air (J/kg to kJ/kg)
            case "Sda": // Specific entropy of dry air (J/kg/K to kJ/kg/K)
            case "Sha": // Specific entropy of humid air (J/kg/K to kJ/kg/K)
            case "S": // Specific entropy of humid air (J/kg/K to kJ/kg/K)
                return value / 1000;
            case "P": // Pressure (Pa to bar)
            case "P_w": // Partial pressure of water (Pa to bar)
                return value / 1e5;
            default:
                return value; // No conversion if unit not found
        }
    }

}
