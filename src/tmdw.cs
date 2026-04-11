using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.Integration;

public static partial class CoolPropWrapper
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

    // Static constructor to initialize dictionaries and load CoolProp
    static CoolPropWrapper()
    {
        InitializeMappings();
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
                continue;
            }
            catch (PathTooLongException)
            {
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
            
            if (string.IsNullOrEmpty(error))
                return "Unknown CoolProp error (no error details available)";
            
            return error.Trim();
        }
        catch (Exception ex)
        {
            return $"Error retrieving CoolProp error: {ex.Message}";
        }
    }
}
