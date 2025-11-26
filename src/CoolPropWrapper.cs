using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.Integration;

public static class CoolPropWrapper
{
    // Static dictionaries for fast lookups (initialized once, thread-safe)
    private static readonly Dictionary<string, string> PropertyNameMap;
    private static readonly Dictionary<string, string> FluidNameMap;
    private static readonly Dictionary<string, string> HumidAirPropertyMap;
    private static readonly HashSet<string> TemperatureProperties;
    private static readonly HashSet<string> PressureProperties;
    private static readonly HashSet<string> EnergyProperties;

    // Static constructor to initialize dictionaries
    static CoolPropWrapper()
    {
        // Initialize property name mappings
        PropertyNameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // Temperature
            ["T"] = "T", ["TEMP"] = "T", ["TEMPERATURE"] = "T",
            // Pressure
            ["P"] = "P", ["PRES"] = "P", ["PRESSURE"] = "P",
            // Enthalpy
            ["H"] = "H", ["ENTH"] = "H", ["ENTHALPY"] = "H", ["HMASS"] = "H", ["H_MASS"] = "H",
            ["HMOLAR"] = "Hmolar", ["H_MOLAR"] = "Hmolar",
            // Internal Energy
            ["U"] = "U", ["INTERNALENERGY"] = "U", ["UMASS"] = "U", ["U_MASS"] = "U",
            ["UMOLAR"] = "Umolar", ["U_MOLAR"] = "Umolar",
            // Entropy
            ["S"] = "S", ["ENTR"] = "S", ["ENTROPY"] = "S", ["SMASS"] = "S", ["S_MASS"] = "S",
            ["SMOLAR"] = "Smolar", ["S_MOLAR"] = "Smolar",
            ["SMOLAR_RESIDUAL"] = "Smolar_residual", ["SMOLARRESIDUAL"] = "Smolar_residual",
            // Density
            ["D"] = "D", ["RHO"] = "D", ["DENS"] = "D", ["DENSITY"] = "D", ["DMASS"] = "D", ["D_MASS"] = "D", ["RHOMASS"] = "D",
            ["DMOLAR"] = "Dmolar", ["RHOMOLAR"] = "Dmolar", ["DMOL"] = "Dmolar", ["D_MOL"] = "Dmolar", ["D_MOLAR"] = "Dmolar",
            // Specific Heat Capacity
            ["CVMASS"] = "Cvmass", ["CV"] = "Cvmass",
            ["CPMASS"] = "Cpmass", ["CP"] = "Cpmass", ["C"] = "Cpmass",
            ["CPMOLAR"] = "Cpmolar", ["CPMOL"] = "Cpmolar",
            ["CVMOLAR"] = "Cvmolar", ["CVMOL"] = "Cvmolar",
            ["CP0MASS"] = "Cp0mass", ["CP0MOLAR"] = "Cp0molar",
            // Quality
            ["Q"] = "Q", ["QUALITY"] = "Q", ["X"] = "Q", ["VAPORFRACTION"] = "Q", ["VAPOR_FRACTION"] = "Q",
            // Reduced properties
            ["TAU"] = "Tau", ["DELTA"] = "Delta",
            // Helmholtz derivatives
            ["ALPHA0"] = "Alpha0", ["ALPHAR"] = "Alphar",
            ["DALPHA0_DDELTA_CONSTTAU"] = "DALPHA0_DDELTA_CONSTTAU", ["DALPHA0_DDELTA"] = "DALPHA0_DDELTA_CONSTTAU",
            ["DALPHA0_DTAU_CONSTDELTA"] = "DALPHA0_DTAU_CONSTDELTA", ["DALPHA0_DTAU"] = "DALPHA0_DTAU_CONSTDELTA",
            ["DALPHAR_DDELTA_CONSTTAU"] = "DALPHAR_DDELTA_CONSTTAU", ["DALPHAR_DDELTA"] = "DALPHAR_DDELTA_CONSTTAU",
            ["DALPHAR_DTAU_CONSTDELTA"] = "DALPHAR_DTAU_CONSTDELTA", ["DALPHAR_DTAU"] = "DALPHAR_DTAU_CONSTDELTA",
            // Speed of sound
            ["SPEED_OF_SOUND"] = "A", ["SPEEDOFSOUND"] = "A", ["A"] = "A", ["W"] = "A",
            // Virial coefficients
            ["BVIRIAL"] = "Bvirial", ["CVIRIAL"] = "Cvirial",
            ["DBVIRIAL_DT"] = "DBVIRIAL_DT", ["DBVIRIAL"] = "DBVIRIAL_DT", ["DBVIRIALDT"] = "DBVIRIAL_DT",
            ["DCVIRIAL_DT"] = "DCVIRIAL_DT", ["DCVIRIAL"] = "DCVIRIAL_DT", ["DCVIRIALDT"] = "DCVIRIAL_DT",
            // Conductivity
            ["K"] = "conductivity", ["CONDUCTIVITY"] = "conductivity", ["L"] = "conductivity",
            // Dipole moment
            ["DIPOLE_MOMENT"] = "DIPOLE_MOMENT", ["DIPOLEMOMENT"] = "DIPOLE_MOMENT", ["DIPOLE"] = "DIPOLE_MOMENT",
            // Gibbs energy
            ["G"] = "G", ["GMASS"] = "G", ["GIBBS"] = "G",
            ["GMOLAR"] = "Gmolar",
            // Helmholtz energy
            ["HELMHOLTZMASS"] = "HELMHOLTZMASS", ["HELMHOLTZMOLAR"] = "HELMHOLTZMOLAR", ["HH"] = "HELMHOLTZMASS",
            // Expansion coefficients
            ["GAMMA"] = "isentropic_expansion_coefficient", ["ISENTROPICEXPANSIONCOEFFICIENT"] = "isentropic_expansion_coefficient",
            ["ISOBARIC_EXPANSION_COEFFICIENT"] = "isobaric_expansion_coefficient", ["ISOBARICEXPANSION"] = "isobaric_expansion_coefficient", ["ISOBARICEXPANSIONCOEFFICIENT"] = "isobaric_expansion_coefficient",
            ["ISOTHERMAL_COMPRESSIBILITY"] = "isothermal_compressibility", ["ISOTHERMALCOMPRESSIBILITY"] = "isothermal_compressibility",
            // Surface tension
            ["SURFACE_TENSION"] = "surface_tension", ["SURFACETENSION"] = "surface_tension", ["SIGMA"] = "surface_tension", ["I"] = "surface_tension",
            // Molar mass
            ["MM"] = "M", ["MOLAR_MASS"] = "M", ["MOLARMASS"] = "M", ["MOLEMASS"] = "M",
            // Critical properties
            ["PCRIT"] = "Pcrit", ["P_CRITICAL"] = "Pcrit", ["PCRITICAL"] = "Pcrit",
            ["TCRIT"] = "Tcrit", ["T_CRITICAL"] = "Tcrit", ["TCRITICAL"] = "Tcrit",
            ["RHOCRIT"] = "rhocrit", ["RHOCRITICAL"] = "rhocrit", ["DCRIT"] = "rhocrit",
            ["RHOMASS_CRITICAL"] = "rhomass_critical", ["RHOMASSCRITICAL"] = "rhomass_critical",
            ["RHOMOLAR_CRITICAL"] = "rhomolar_critical", ["RHOMOLARCRITICAL"] = "rhomolar_critical",
            // Phase
            ["PHASE"] = "Phase",
            // Max/Min properties
            ["PMAX"] = "pmax", ["P_MAX"] = "pmax",
            ["PMIN"] = "pmin", ["P_MIN"] = "pmin",
            ["TMAX"] = "Tmax", ["T_MAX"] = "Tmax",
            ["TMIN"] = "Tmin", ["T_MIN"] = "Tmin",
            // Prandtl number
            ["PRANDTL"] = "Prandtl",
            // Triple point
            ["PTRIPLE"] = "ptriple", ["P_TRIPLE"] = "ptriple", ["PTRIP"] = "ptriple",
            ["TTRIPLE"] = "Ttriple", ["T_TRIPLE"] = "Ttriple", ["TTRIP"] = "Ttriple",
            // Reducing properties
            ["P_REDUCING"] = "p_reducing", ["PREDUCING"] = "p_reducing",
            ["T_REDUCING"] = "T_reducing", ["TREDUCING"] = "T_reducing",
            ["RHOMASS_REDUCING"] = "rhomass_reducing", ["RHOMASSREDUCING"] = "rhomass_reducing",
            ["RHOMOLAR_REDUCING"] = "rhomolar_reducing", ["RHOMOLARREDUCING"] = "rhomolar_reducing",
            // Freeze temperature
            ["T_FREEZE"] = "T_freeze", ["TFREEZE"] = "T_freeze", ["FREEZING_TEMPERATURE"] = "T_freeze",
            // Viscosity
            ["MU"] = "viscosity", ["VISCOSITY"] = "viscosity", ["V"] = "viscosity",
            // Compressibility factor
            ["Z"] = "Z", ["COMPRESSIBILITY"] = "Z", ["COMPRESSIBILITYFACTOR"] = "Z",
            // Acentric factor
            ["ACENTRIC"] = "acentric", ["ACENTRIC_FACTOR"] = "acentric", ["ACENTRICFACTOR"] = "acentric", ["OMEGA"] = "acentric",
            // Other properties
            ["FUNDAMENTAL_DERIVATIVE_OF_GAS_DYNAMICS"] = "fundamental_derivative_of_gas_dynamics", ["FUNDAMENTALDERIVATIVE"] = "fundamental_derivative_of_gas_dynamics", ["FH"] = "fundamental_derivative_of_gas_dynamics",
            ["GAS_CONSTANT"] = "gas_constant", ["GASCONSTANT"] = "gas_constant",
            ["FRACTION_MAX"] = "fraction_max", ["FRACTIONMAX"] = "fraction_max",
            ["FRACTION_MIN"] = "fraction_min", ["FRACTIONMIN"] = "fraction_min",
            ["GWP20"] = "GWP20", ["GWP_20"] = "GWP20",
            ["GWP100"] = "GWP100", ["GWP_100"] = "GWP100",
            ["GWP500"] = "GWP500", ["GWP_500"] = "GWP500",
            ["ODP"] = "ODP", ["OZONEDEPLETIONPOTENTIAL"] = "ODP",
            ["PIP"] = "PIP"
        };

        // Initialize fluid name mappings
        FluidNameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["1BUTENE"] = "1-Butene", ["ACETONE"] = "Acetone", ["AIR"] = "Air",
            ["AMMONIA"] = "Ammonia", ["NH3"] = "Ammonia",
            ["ARGON"] = "Argon", ["AR"] = "Argon",
            ["BENZENE"] = "Benzene",
            ["CARBONDIOXIDE"] = "CarbonDioxide", ["CO2"] = "CarbonDioxide",
            ["CARBONMONOXIDE"] = "CarbonMonoxide", ["CO"] = "CarbonMonoxide",
            ["CARBONYLSULFIDE"] = "CarbonylSulfide", ["COS"] = "CarbonylSulfide",
            ["CIS2BUTENE"] = "cis-2-Butene",
            ["CYCLOHEXANE"] = "CycloHexane", ["CYCLOPENTANE"] = "Cyclopentane", ["CYCLOPROPANE"] = "CycloPropane",
            ["D4"] = "D4", ["D5"] = "D5", ["D6"] = "D6",
            ["DEUTERIUM"] = "Deuterium", ["D2"] = "Deuterium",
            ["DICHLOROETHANE"] = "Dichloroethane", ["DIETHYLETHER"] = "DiethylEther",
            ["DIMETHYLCARBONATE"] = "DimethylCarbonate",
            ["DIMETHYLETHER"] = "DimethylEther", ["DME"] = "DimethylEther",
            ["ETHANE"] = "Ethane", ["C2H6"] = "Ethane",
            ["ETHANOL"] = "Ethanol", ["ETHYLBENZENE"] = "EthylBenzene",
            ["ETHYLENE"] = "Ethylene", ["C2H4"] = "Ethylene",
            ["ETHYLENEOXIDE"] = "EthyleneOxide",
            ["FLUORINE"] = "Fluorine", ["F2"] = "Fluorine",
            ["HEAVYWATER"] = "HeavyWater", ["D2O"] = "HeavyWater",
            ["HELIUM"] = "Helium", ["HE"] = "Helium",
            ["HFE143M"] = "HFE143m",
            ["HYDROGEN"] = "Hydrogen", ["H2"] = "Hydrogen",
            ["HYDROGENCHLORIDE"] = "HydrogenChloride", ["HCL"] = "HydrogenChloride",
            ["HYDROGENSULFIDE"] = "HydrogenSulfide", ["H2S"] = "HydrogenSulfide",
            ["ISOBUTANE"] = "IsoButane", ["IBUTANE"] = "IsoButane",
            ["ISOBUTENE"] = "IsoButene", ["IBUTENE"] = "IsoButene",
            ["ISOHEXANE"] = "Isohexane", ["ISOPENTANE"] = "Isopentane",
            ["KRYPTON"] = "Krypton", ["KR"] = "Krypton",
            ["MXYLENE"] = "m-Xylene",
            ["MD2M"] = "MD2M", ["MD3M"] = "MD3M", ["MD4M"] = "MD4M", ["MDM"] = "MDM",
            ["METHANE"] = "Methane", ["CH4"] = "Methane",
            ["METHANOL"] = "Methanol", ["MEOH"] = "Methanol",
            ["METHYLLINOLEATE"] = "MethylLinoleate", ["METHYLLINOLENATE"] = "MethylLinolenate",
            ["METHYLOLEATE"] = "MethylOleate", ["METHYLPALMITATE"] = "MethylPalmitate",
            ["METHYLSTEARATE"] = "MethylStearate",
            ["MM"] = "MM",
            ["NBUTANE"] = "n-Butane", ["BUTANE"] = "n-Butane",
            ["NDECANE"] = "n-Decane", ["DECANE"] = "n-Decane",
            ["NDODECANE"] = "n-Dodecane", ["DODECANE"] = "n-Dodecane",
            ["NHEPTANE"] = "n-Heptane", ["HEPTANE"] = "n-Heptane",
            ["NHEXANE"] = "n-Hexane", ["HEXANE"] = "n-Hexane",
            ["NNONANE"] = "n-Nonane", ["NONANE"] = "n-Nonane",
            ["NOCTANE"] = "n-Octane", ["OCTANE"] = "n-Octane",
            ["NPENTANE"] = "n-Pentane", ["PENTANE"] = "n-Pentane",
            ["NPROPANE"] = "n-Propane", ["PROPANE"] = "n-Propane",
            ["NUNDECANE"] = "n-Undecane", ["UNDECANE"] = "n-Undecane",
            ["NEON"] = "Neon", ["NE"] = "Neon",
            ["NEOPENTANE"] = "Neopentane",
            ["NITROGEN"] = "Nitrogen", ["N2"] = "Nitrogen",
            ["NITROUSOXIDE"] = "NitrousOxide", ["N2O"] = "NitrousOxide",
            ["NOVEC649"] = "Novec649",
            ["OXYLENE"] = "o-Xylene",
            ["ORTHODEUTERIUM"] = "OrthoDeuterium", ["ORTHOHYDROGEN"] = "OrthoHydrogen",
            ["OXYGEN"] = "Oxygen", ["O2"] = "Oxygen",
            ["PXYLENE"] = "p-Xylene",
            ["PARADEUTERIUM"] = "ParaDeuterium", ["PARAHYDROGEN"] = "ParaHydrogen",
            ["PROPYLENE"] = "Propylene", ["C3H6"] = "Propylene",
            ["PROPYNE"] = "Propyne",
            // Refrigerants
            ["R11"] = "R11", ["R113"] = "R113", ["R114"] = "R114", ["R115"] = "R115",
            ["R116"] = "R116", ["R12"] = "R12", ["R123"] = "R123",
            ["R1233ZDE"] = "R1233zd(E)", ["R1233ZD(E)"] = "R1233zd(E)",
            ["R1234YF"] = "R1234yf",
            ["R1234ZEE"] = "R1234ze(E)", ["R1234ZE(E)"] = "R1234ze(E)",
            ["R1234ZEZ"] = "R1234ze(Z)", ["R1234ZE(Z)"] = "R1234ze(Z)",
            ["R124"] = "R124", ["R125"] = "R125", ["R13"] = "R13",
            ["R134A"] = "R134a", ["R13I1"] = "R13I1", ["R14"] = "R14",
            ["R141B"] = "R141b", ["R142B"] = "R142b", ["R143A"] = "R143a",
            ["R152A"] = "R152A", ["R161"] = "R161", ["R21"] = "R21",
            ["R218"] = "R218", ["R22"] = "R22", ["R227EA"] = "R227EA",
            ["R23"] = "R23", ["R236EA"] = "R236EA", ["R236FA"] = "R236FA",
            ["R245CA"] = "R245ca", ["R245FA"] = "R245fa", ["R32"] = "R32",
            ["R365MFC"] = "R365MFC", ["R40"] = "R40",
            ["R404A"] = "R404A", ["R407C"] = "R407C", ["R41"] = "R41",
            ["R410A"] = "R410A", ["R507A"] = "R507A",
            ["RC318"] = "RC318",
            ["SES36"] = "SES36",
            ["SULFURDIOXIDE"] = "SulfurDioxide", ["SO2"] = "SulfurDioxide",
            ["SULFURHEXAFLUORIDE"] = "SulfurHexafluoride", ["SF6"] = "SulfurHexafluoride",
            ["TOLUENE"] = "Toluene",
            ["TRANS2BUTENE"] = "trans-2-Butene",
            ["WATER"] = "Water", ["H2O"] = "Water",
            ["XENON"] = "Xenon", ["XE"] = "Xenon"
        };

        // Initialize humid air property mappings
        HumidAirPropertyMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["B"] = "Twb", ["TWB"] = "Twb", ["T_WB"] = "Twb", ["WETBULB"] = "Twb", ["WETBULBTEMP"] = "Twb", ["WETBULBTEMPERATURE"] = "Twb",
            ["C"] = "Cda", ["CP"] = "Cda", ["CDA"] = "Cda", ["CPDA"] = "Cda", ["CP_DA"] = "Cda",
            ["CHA"] = "Cha", ["CPHA"] = "Cha", ["CP_HA"] = "Cha",
            ["CV"] = "CV", ["CVMASS"] = "CV",
            ["CVHA"] = "CVha", ["CV_HA"] = "CVha",
            ["D"] = "Tdp", ["TDP"] = "Tdp", ["T_DP"] = "Tdp", ["DEWPOINT"] = "Tdp", ["DEWPOINTTEMP"] = "Tdp", ["DEWPOINTTEMPERATURE"] = "Tdp",
            ["H"] = "Hda", ["HDA"] = "Hda", ["H_DA"] = "Hda", ["ENTHALPY"] = "Hda",
            ["HHA"] = "Hha", ["H_HA"] = "Hha",
            ["K"] = "K", ["CONDUCTIVITY"] = "K", ["THERMALCONDUCTIVITY"] = "K",
            ["M"] = "MU", ["MU"] = "MU", ["VISC"] = "MU", ["VISCOSITY"] = "MU", ["DYNAMICVISCOSITY"] = "MU",
            ["PSI_W"] = "Psi_w", ["PSIW"] = "Psi_w", ["Y"] = "Psi_w",
            ["P"] = "P", ["PRESSURE"] = "P", ["PRES"] = "P",
            ["P_W"] = "P_w", ["PW"] = "P_w", ["PARTIALPRESSURE"] = "P_w", ["WATERPRESSURE"] = "P_w",
            ["R"] = "R", ["RH"] = "R", ["RELHUM"] = "R", ["RELATIVEHUMIDITY"] = "R",
            ["S"] = "Sda", ["SDA"] = "Sda", ["S_DA"] = "Sda", ["ENTROPY"] = "Sda",
            ["SHA"] = "Sha", ["S_HA"] = "Sha",
            ["T"] = "T", ["TDB"] = "T", ["T_DB"] = "T", ["TEMP"] = "T", ["TEMPERATURE"] = "T", ["DRYBULB"] = "T", ["DRYBULBTEMP"] = "T", ["DRYBULBTEMPERATURE"] = "T",
            ["V"] = "Vda", ["VDA"] = "Vda", ["V_DA"] = "Vda",
            ["VHA"] = "Vha", ["V_HA"] = "Vha",
            ["W"] = "W", ["OMEGA"] = "W", ["HUMRAT"] = "W", ["HUMIDITYRATIO"] = "W", ["MIXINGRATIO"] = "W",
            ["Z"] = "Z", ["COMPRESSIBILITY"] = "Z", ["COMPRESSIBILITYFACTOR"] = "Z"
        };

        // Properties requiring temperature conversion (°C <-> K)
        TemperatureProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "T", "Tcrit", "Tmax", "Tmin", "Ttriple", "T_freeze", "T_reducing", "Twb", "Tdp"
        };

        // Properties requiring pressure conversion (bar <-> Pa)
        PressureProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "P", "Pcrit", "pmax", "pmin", "ptriple", "p_reducing", "P_w"
        };

        // Properties requiring energy conversion (kJ <-> J)
        EnergyProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "H", "Hmolar", "U", "Umolar", "S", "Smolar", "Smolar_residual",
            "Cpmass", "Cvmass", "Cpmolar", "Cvmolar", "Cp0mass", "Cp0molar",
            "G", "Gmolar", "Hda", "Hha", "Sda", "Sha", "Cda", "Cha"
        };

        LoadCoolPropDll();
    }
    // Windows API for LoadLibrary and SetDllDirectory
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr LoadLibrary(string dllToLoad);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GetProcAddress(IntPtr hModule, string procedureName);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool SetDllDirectory(string lpPathName);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool AddDllDirectory(string NewDirectory);

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
    [ExcelFunction(Name = "PropsSI", Description = "Calculate thermodynamic properties of real fluids using CoolProp with SI units (K, Pa, J/kg, etc.). Supports array inputs for value1 and/or value2.")]
    public static object PropsSI(string output, string name1, object value1, string name2, object value2, string fluid)
    {
        // Check for missing or invalid inputs (strings)
        if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
        if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
        if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
        if (string.IsNullOrWhiteSpace(fluid)) return "Error: Fluid name is missing.";
        if (value1 == null) return "Error: First property value is missing.";
        if (value2 == null) return "Error: Second property value is missing.";

        // Normalize parameter names and fluid name
        name1 = FormatName(name1);
        name2 = FormatName(name2);
        output = FormatName(output);
        fluid = FormatFluidName(fluid);

        // Check if either input is an array
        bool isValue1Array = IsArrayInput(value1);
        bool isValue2Array = IsArrayInput(value2);

        // If neither is an array, handle as a single value (original behavior)
        if (!isValue1Array && !isValue2Array)
        {
            if (!(value1 is double)) return "Error: First property value is not a number.";
            if (!(value2 is double)) return "Error: Second property value is not a number.";

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

            // Check if the result indicates an error
            if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
            {
                return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
            }

            // Return result in SI units (no conversion)
            return result;
        }

        // Handle array inputs
        try
        {
            double[] values1;
            double[] values2;

            // Extract arrays or convert single values to arrays
            if (isValue1Array)
            {
                values1 = ExtractDoubleArray(value1);
            }
            else
            {
                if (!(value1 is double)) return "Error: First property value is not a number.";
                values1 = new double[] { (double)value1 };
            }

            if (isValue2Array)
            {
                values2 = ExtractDoubleArray(value2);
            }
            else
            {
                if (!(value2 is double)) return "Error: Second property value is not a number.";
                values2 = new double[] { (double)value2 };
            }

            // Check if both are arrays and have the same length
            if (isValue1Array && isValue2Array && values1.Length != values2.Length)
            {
                return $"Error: Array lengths must match. value1 has {values1.Length} elements, value2 has {values2.Length} elements.";
            }

            // Determine the result array length
            int resultLength = Math.Max(values1.Length, values2.Length);
            double[] results = new double[resultLength];

            // Calculate properties for each pair
            for (int i = 0; i < resultLength; i++)
            {
                double val1 = values1[isValue1Array ? i : 0];
                double val2 = values2[isValue2Array ? i : 0];

                // Call CoolProp directly (SI units, no conversion)
                double result;
                try
                {
                    result = PropsSI_Internal(output, name1, val1, name2, val2, fluid);
                }
                catch (Exception ex)
                {
                    return $"Error at index {i}: {ex.Message}. CoolProp error: {GetCoolPropError()}";
                }

                // Check if the result indicates an error
                if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
                {
                    return $"Error at index {i}: CoolProp failed to compute property. {GetCoolPropError()}";
                }

                results[i] = result;
            }

            // Determine output orientation based on input orientation
            bool outputAsRow = (isValue1Array && IsRowArray(value1)) || (isValue2Array && IsRowArray(value2));
            
            object[,] outputArray;
            if (outputAsRow)
            {
                // Return as a row array (1 row, multiple columns)
                outputArray = new object[1, resultLength];
                for (int i = 0; i < resultLength; i++)
                {
                    outputArray[0, i] = results[i];
                }
            }
            else
            {
                // Return as a column array (multiple rows, 1 column)
                outputArray = new object[resultLength, 1];
                for (int i = 0; i < resultLength; i++)
                {
                    outputArray[i, 0] = results[i];
                }
            }
            return outputArray;
        }
        catch (InvalidCastException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (ArgumentException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (Exception ex)
        {
            return $"Error processing array inputs: {ex.Message}";
        }
    }

    // Function to calculate thermodynamic properties using engineering units
    [ExcelFunction(Name = "Props", Description = "Calculate thermodynamic properties of real fluids using CoolProp with engineering units (°C, bar, kJ/kg, etc.). Supports array inputs for value1 and/or value2.")]
    public static object Props(string output, string name1, object value1, string name2, object value2, string fluid)
    {
        // Check for missing or invalid inputs (strings)
        if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
        if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
        if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
        if (string.IsNullOrWhiteSpace(fluid)) return "Error: Fluid name is missing.";
        if (value1 == null) return "Error: First property value is missing.";
        if (value2 == null) return "Error: Second property value is missing.";

        // Normalize parameter names and fluid name
        name1 = FormatName(name1);
        name2 = FormatName(name2);
        output = FormatName(output);
        fluid = FormatFluidName(fluid);

        // Check if either input is an array
        bool isValue1Array = IsArrayInput(value1);
        bool isValue2Array = IsArrayInput(value2);

        // If neither is an array, handle as a single value (original behavior)
        if (!isValue1Array && !isValue2Array)
        {
            if (!(value1 is double)) return "Error: First property value is not a number.";
            if (!(value2 is double)) return "Error: Second property value is not a number.";

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

            // Check if the result indicates an error
            if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
            {
                return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
            }

            // Convert output to engineering units
            return ConvertFromSI(output, result);
        }

        // Handle array inputs
        try
        {
            double[] values1;
            double[] values2;

            // Extract arrays or convert single values to arrays
            if (isValue1Array)
            {
                values1 = ExtractDoubleArray(value1);
            }
            else
            {
                if (!(value1 is double)) return "Error: First property value is not a number.";
                values1 = new double[] { (double)value1 };
            }

            if (isValue2Array)
            {
                values2 = ExtractDoubleArray(value2);
            }
            else
            {
                if (!(value2 is double)) return "Error: Second property value is not a number.";
                values2 = new double[] { (double)value2 };
            }

            // Check if both are arrays and have the same length
            if (isValue1Array && isValue2Array && values1.Length != values2.Length)
            {
                return $"Error: Array lengths must match. value1 has {values1.Length} elements, value2 has {values2.Length} elements.";
            }

            // Determine the result array length
            int resultLength = Math.Max(values1.Length, values2.Length);
            double[] results = new double[resultLength];

            // Calculate properties for each pair
            for (int i = 0; i < resultLength; i++)
            {
                double val1 = values1[isValue1Array ? i : 0];
                double val2 = values2[isValue2Array ? i : 0];

                // Convert inputs to SI units
                double val1SI = ConvertToSI(name1, val1);
                double val2SI = ConvertToSI(name2, val2);

                // Call CoolProp for the requested property
                double result;
                try
                {
                    result = PropsSI_Internal(output, name1, val1SI, name2, val2SI, fluid);
                }
                catch (Exception ex)
                {
                    return $"Error at index {i}: {ex.Message}. CoolProp error: {GetCoolPropError()}";
                }

                // Check if the result indicates an error
                if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
                {
                    return $"Error at index {i}: CoolProp failed to compute property. {GetCoolPropError()}";
                }

                // Convert output to engineering units
                results[i] = ConvertFromSI(output, result);
            }

            // Determine output orientation based on input orientation
            bool outputAsRow = (isValue1Array && IsRowArray(value1)) || (isValue2Array && IsRowArray(value2));
            
            object[,] outputArray;
            if (outputAsRow)
            {
                // Return as a row array (1 row, multiple columns)
                outputArray = new object[1, resultLength];
                for (int i = 0; i < resultLength; i++)
                {
                    outputArray[0, i] = results[i];
                }
            }
            else
            {
                // Return as a column array (multiple rows, 1 column)
                outputArray = new object[resultLength, 1];
                for (int i = 0; i < resultLength; i++)
                {
                    outputArray[i, 0] = results[i];
                }
            }
            return outputArray;
        }
        catch (InvalidCastException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (ArgumentException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (Exception ex)
        {
            return $"Error processing array inputs: {ex.Message}";
        }
    }

    // TMPr alias for Props (uses engineering units)
    [ExcelFunction(Name = "TMPr", Description = "Calculate thermodynamic properties of real fluids using CoolProp with engineering units (°C, bar, kJ/kg, etc.). Alias for Props.")]
    public static object TMPr(string output, string name1, object value1, string name2, object value2, string fluid)
    {
        return Props(output, name1, value1, name2, value2, fluid);
    }

    [ExcelFunction(Name = "HAPropsSI", Description = "Calculate thermodynamic properties of humid air using CoolProp with SI units (K, Pa, J/kg, etc.). Supports array inputs.")]
    public static object HAPropsSI(string output, string name1, object value1, string name2, object value2, string name3, object value3)
    {
        // Check for missing or invalid inputs (strings)
        if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
        if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
        if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
        if (string.IsNullOrWhiteSpace(name3)) return "Error: Third property name is missing.";
        if (value1 == null) return "Error: First property value is missing.";
        if (value2 == null) return "Error: Second property value is missing.";
        if (value3 == null) return "Error: Third property value is missing.";

        // Normalize parameter names
        name1 = FormatName_HA(name1);
        name2 = FormatName_HA(name2);
        name3 = FormatName_HA(name3);
        output = FormatName_HA(output);

        // Check if any input is an array
        bool isValue1Array = IsArrayInput(value1);
        bool isValue2Array = IsArrayInput(value2);
        bool isValue3Array = IsArrayInput(value3);

        // If none is an array, handle as a single value (original behavior)
        if (!isValue1Array && !isValue2Array && !isValue3Array)
        {
            if (!(value1 is double)) return "Error: First property value is not a number.";
            if (!(value2 is double)) return "Error: Second property value is not a number.";
            if (!(value3 is double)) return "Error: Third property value is not a number.";

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

            // Check if the result indicates an error
            if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
            {
                return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
            }

            // Return result in SI units (no conversion)
            return result;
        }

        // Handle array inputs
        try
        {
            double[] values1, values2, values3;

            // Extract arrays or convert single values to arrays
            if (isValue1Array)
                values1 = ExtractDoubleArray(value1);
            else
            {
                if (!(value1 is double)) return "Error: First property value is not a number.";
                values1 = new double[] { (double)value1 };
            }

            if (isValue2Array)
                values2 = ExtractDoubleArray(value2);
            else
            {
                if (!(value2 is double)) return "Error: Second property value is not a number.";
                values2 = new double[] { (double)value2 };
            }

            if (isValue3Array)
                values3 = ExtractDoubleArray(value3);
            else
            {
                if (!(value3 is double)) return "Error: Third property value is not a number.";
                values3 = new double[] { (double)value3 };
            }

            // Check if all arrays have the same length
            int resultLength = Math.Max(Math.Max(values1.Length, values2.Length), values3.Length);
            if ((isValue1Array && values1.Length != resultLength && values1.Length > 1) ||
                (isValue2Array && values2.Length != resultLength && values2.Length > 1) ||
                (isValue3Array && values3.Length != resultLength && values3.Length > 1))
            {
                return $"Error: All array inputs must have the same length. Lengths: value1={values1.Length}, value2={values2.Length}, value3={values3.Length}";
            }

            double[] results = new double[resultLength];

            // Calculate properties for each set
            for (int i = 0; i < resultLength; i++)
            {
                double val1 = values1[isValue1Array ? Math.Min(i, values1.Length - 1) : 0];
                double val2 = values2[isValue2Array ? Math.Min(i, values2.Length - 1) : 0];
                double val3 = values3[isValue3Array ? Math.Min(i, values3.Length - 1) : 0];

                // Call CoolProp directly (SI units, no conversion)
                double result;
                try
                {
                    result = HAPropsSI_Internal(output, name1, val1, name2, val2, name3, val3);
                }
                catch (Exception ex)
                {
                    return $"Error at index {i}: {ex.Message}. CoolProp error: {GetCoolPropError()}";
                }

                // Check if the result indicates an error
                if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
                {
                    return $"Error at index {i}: CoolProp failed to compute property. {GetCoolPropError()}";
                }

                results[i] = result;
            }

            // Determine output orientation based on input orientation
            bool outputAsRow = (isValue1Array && IsRowArray(value1)) || 
                              (isValue2Array && IsRowArray(value2)) || 
                              (isValue3Array && IsRowArray(value3));
            
            object[,] outputArray;
            if (outputAsRow)
            {
                // Return as a row array (1 row, multiple columns)
                outputArray = new object[1, resultLength];
                for (int i = 0; i < resultLength; i++)
                {
                    outputArray[0, i] = results[i];
                }
            }
            else
            {
                // Return as a column array (multiple rows, 1 column)
                outputArray = new object[resultLength, 1];
                for (int i = 0; i < resultLength; i++)
                {
                    outputArray[i, 0] = results[i];
                }
            }
            return outputArray;
        }
        catch (InvalidCastException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (ArgumentException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (Exception ex)
        {
            return $"Error processing array inputs: {ex.Message}";
        }
    }

    [ExcelFunction(Name = "HAProps", Description = "Calculate thermodynamic properties of humid air using CoolProp with engineering units (°C, bar, kJ/kg, etc.). Supports array inputs.")]
    public static object HAProps(string output, string name1, object value1, string name2, object value2, string name3, object value3)
    {
        // Check for missing or invalid inputs (strings)
        if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
        if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
        if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
        if (string.IsNullOrWhiteSpace(name3)) return "Error: Third property name is missing.";
        if (value1 == null) return "Error: First property value is missing.";
        if (value2 == null) return "Error: Second property value is missing.";
        if (value3 == null) return "Error: Third property value is missing.";

        // Normalize parameter names
        name1 = FormatName_HA(name1);
        name2 = FormatName_HA(name2);
        name3 = FormatName_HA(name3);
        output = FormatName_HA(output);

        // Check if any input is an array
        bool isValue1Array = IsArrayInput(value1);
        bool isValue2Array = IsArrayInput(value2);
        bool isValue3Array = IsArrayInput(value3);

        // If none is an array, handle as a single value (original behavior)
        if (!isValue1Array && !isValue2Array && !isValue3Array)
        {
            if (!(value1 is double)) return "Error: First property value is not a number.";
            if (!(value2 is double)) return "Error: Second property value is not a number.";
            if (!(value3 is double)) return "Error: Third property value is not a number.";

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

            // Check if the result indicates an error
            if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
            {
                return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
            }

            // Convert output to engineering units
            return ConvertFromSI_HA(output, result);
        }

        // Handle array inputs
        try
        {
            double[] values1, values2, values3;

            // Extract arrays or convert single values to arrays
            if (isValue1Array)
                values1 = ExtractDoubleArray(value1);
            else
            {
                if (!(value1 is double)) return "Error: First property value is not a number.";
                values1 = new double[] { (double)value1 };
            }

            if (isValue2Array)
                values2 = ExtractDoubleArray(value2);
            else
            {
                if (!(value2 is double)) return "Error: Second property value is not a number.";
                values2 = new double[] { (double)value2 };
            }

            if (isValue3Array)
                values3 = ExtractDoubleArray(value3);
            else
            {
                if (!(value3 is double)) return "Error: Third property value is not a number.";
                values3 = new double[] { (double)value3 };
            }

            // Check if all arrays have the same length
            int resultLength = Math.Max(Math.Max(values1.Length, values2.Length), values3.Length);
            if ((isValue1Array && values1.Length != resultLength && values1.Length > 1) ||
                (isValue2Array && values2.Length != resultLength && values2.Length > 1) ||
                (isValue3Array && values3.Length != resultLength && values3.Length > 1))
            {
                return $"Error: All array inputs must have the same length. Lengths: value1={values1.Length}, value2={values2.Length}, value3={values3.Length}";
            }

            double[] results = new double[resultLength];

            // Calculate properties for each set
            for (int i = 0; i < resultLength; i++)
            {
                double val1 = values1[isValue1Array ? Math.Min(i, values1.Length - 1) : 0];
                double val2 = values2[isValue2Array ? Math.Min(i, values2.Length - 1) : 0];
                double val3 = values3[isValue3Array ? Math.Min(i, values3.Length - 1) : 0];

                // Convert inputs to SI units
                double val1SI = ConvertToSI_HA(name1, val1);
                double val2SI = ConvertToSI_HA(name2, val2);
                double val3SI = ConvertToSI_HA(name3, val3);

                // Call CoolProp for the requested property
                double result;
                try
                {
                    result = HAPropsSI_Internal(output, name1, val1SI, name2, val2SI, name3, val3SI);
                }
                catch (Exception ex)
                {
                    return $"Error at index {i}: {ex.Message}. CoolProp error: {GetCoolPropError()}";
                }

                // Check if the result indicates an error
                if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
                {
                    return $"Error at index {i}: CoolProp failed to compute property. {GetCoolPropError()}";
                }

                // Convert output to engineering units
                results[i] = ConvertFromSI_HA(output, result);
            }

            // Determine output orientation based on input orientation
            bool outputAsRow = (isValue1Array && IsRowArray(value1)) || 
                              (isValue2Array && IsRowArray(value2)) || 
                              (isValue3Array && IsRowArray(value3));
            
            object[,] outputArray;
            if (outputAsRow)
            {
                // Return as a row array (1 row, multiple columns)
                outputArray = new object[1, resultLength];
                for (int i = 0; i < resultLength; i++)
                {
                    outputArray[0, i] = results[i];
                }
            }
            else
            {
                // Return as a column array (multiple rows, 1 column)
                outputArray = new object[resultLength, 1];
                for (int i = 0; i < resultLength; i++)
                {
                    outputArray[i, 0] = results[i];
                }
            }
            return outputArray;
        }
        catch (InvalidCastException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (ArgumentException ex)
        {
            return $"Error: {ex.Message}";
        }
        catch (Exception ex)
        {
            return $"Error processing array inputs: {ex.Message}";
        }
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

    // Helper functions for array detection and extraction
    private static bool IsArrayInput(object input)
    {
        return input is object[,] || input is double[,];
    }

    private static bool IsRowArray(object input)
    {
        if (input is object[,] arr2D)
        {
            return arr2D.GetLength(0) == 1 && arr2D.GetLength(1) > 1;
        }
        else if (input is double[,] dblArr2D)
        {
            return dblArr2D.GetLength(0) == 1 && dblArr2D.GetLength(1) > 1;
        }
        return false;
    }

    private static double[] ExtractDoubleArray(object input)
    {
        if (input is object[,] arr2D)
        {
            // Determine if it's a row or column
            int rows = arr2D.GetLength(0);
            int cols = arr2D.GetLength(1);
            
            if (rows == 1) // Row array
            {
                var result = new double[cols];
                for (int i = 0; i < cols; i++)
                {
                    if (arr2D[0, i] is double d)
                        result[i] = d;
                    else
                        throw new InvalidCastException($"Element at position {i} is not a number.");
                }
                return result;
            }
            else if (cols == 1) // Column array
            {
                var result = new double[rows];
                for (int i = 0; i < rows; i++)
                {
                    if (arr2D[i, 0] is double d)
                        result[i] = d;
                    else
                        throw new InvalidCastException($"Element at position {i} is not a number.");
                }
                return result;
            }
            else
            {
                throw new ArgumentException("Array must be a single row or single column.");
            }
        }
        else if (input is double[,] dblArr2D)
        {
            int rows = dblArr2D.GetLength(0);
            int cols = dblArr2D.GetLength(1);
            
            if (rows == 1) // Row array
            {
                var result = new double[cols];
                for (int i = 0; i < cols; i++)
                    result[i] = dblArr2D[0, i];
                return result;
            }
            else if (cols == 1) // Column array
            {
                var result = new double[rows];
                for (int i = 0; i < rows; i++)
                    result[i] = dblArr2D[i, 0];
                return result;
            }
            else
            {
                throw new ArgumentException("Array must be a single row or single column.");
            }
        }
        
        throw new ArgumentException("Input is not a valid array type.");
    }

    // Optimized property name mapping using dictionary lookups
    private static string FormatName(string name)
    {
        if (PropertyNameMap.TryGetValue(name, out string mapped))
            return mapped;
        return name; // Return as is if no match found
    }

    // Optimize fluid name mapping using dictionary lookups
    private static string FormatFluidName(string fluid)
    {
        // Handle mixture strings (contain "&" or "HEOS::")
        if (fluid.Contains("&") || fluid.StartsWith("HEOS::", StringComparison.OrdinalIgnoreCase))
            return fluid;

        // Normalize by removing dashes, underscores, and spaces
        string normalized = fluid.Replace("-", "").Replace("_", "").Replace(" ", "");
        
        if (FluidNameMap.TryGetValue(normalized, out string mapped))
            return mapped;
        
        return fluid; // Return as is if no match found
    }

    // Optimized humid air property name mapping
    private static string FormatName_HA(string name)
    {
        if (HumidAirPropertyMap.TryGetValue(name, out string mapped))
            return mapped;
        return name; // Return as is if no match found
    }

    // Optimized unit conversion to SI using HashSet lookups
    private static double ConvertToSI(string name, double value)
    {
        if (TemperatureProperties.Contains(name))
            return value + 273.15;
        if (PressureProperties.Contains(name))
            return value * 1e5;
        if (EnergyProperties.Contains(name))
            return value * 1000;
        return value; // No conversion if unit not found
    }

    // Optimized unit conversion from SI using HashSet lookups
    private static double ConvertFromSI(string name, double value)
    {
        if (TemperatureProperties.Contains(name))
            return value - 273.15;
        if (PressureProperties.Contains(name))
            return value / 1e5;
        if (EnergyProperties.Contains(name))
            return value / 1000;
        return value; // No conversion if unit not found
    }

    // Convert from custom units to SI units for HA properties
    private static double ConvertToSI_HA(string name, double value)
    {
        if (TemperatureProperties.Contains(name))
            return value + 273.15;
        if (PressureProperties.Contains(name))
            return value * 1e5;
        if (EnergyProperties.Contains(name))
            return value * 1000;
        return value; // No conversion if unit not found
    }

    // Convert from SI units to custom units for HA properties
    private static double ConvertFromSI_HA(string name, double value)
    {
        if (TemperatureProperties.Contains(name))
            return value - 273.15;
        if (PressureProperties.Contains(name))
            return value / 1e5;
        if (EnergyProperties.Contains(name))
            return value / 1000;
        return value; // No conversion if unit not found
    }

}
