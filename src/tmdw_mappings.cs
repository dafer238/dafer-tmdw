using System;
using System.Collections.Generic;

public static partial class CoolPropWrapper
{
    // Static dictionaries for fast lookups (initialized once via static constructor, thread-safe)
    private static Dictionary<string, string> PropertyNameMap;
    private static Dictionary<string, string> FluidNameMap;
    private static Dictionary<string, string> HumidAirPropertyMap;
    private static HashSet<string> TemperatureProperties;
    private static HashSet<string> PressureProperties;
    private static HashSet<string> EnergyProperties;

    private static void InitializeMappings()
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
    }
}
