# A CoolProp wrapper for Excel in engineering units

## Installation

[The latest release](https://github.com/Danisaski/dafer-tmpr/releases) is available for a portable standalone version.

After downloading, extract the zip file with the contents; `Coolprop.dll`, `CoolpropWrapper.dll`, `CoolpropWrapper.xll` and `CoolPropWrapper.dna`, place the four files **together** in the directory of your choice. Finally import in Excel the `.xll` file via: File (Archivo) -> Options (Opciones) -> Addins (Complementos) -> Import (Ir...) -> Browse... (Examinar...) and select the `.xll` file.

## Usage

CoolPropWrapper allows you to compute thermodynamic properties of real fluids and humid air using CoolProp in Excel. The wrapper uses **standard CoolProp function names** with support for both engineering and SI units. All functions support **case-insensitive** parameter names for flexibility.

- **Engineering Units** (default): Temperature in °C, Pressure in bar, Energy in kJ/kg, etc.
- **SI Units**: Temperature in K, Pressure in Pa, Energy in J/kg, etc.

### Real Fluid Functions

#### `Props(output, name1, value1, name2, value2, fluid)`

- Computes thermodynamic properties of real fluids using **engineering units**.
- Standard CoolProp function with engineering units instead of SI.
- **Units used:**
  - Temperature: **Celsius (°C)** (converted internally to Kelvin (K))
  - Pressure: **bar** (converted internally to Pascal (Pa))
  - Enthalpy, Internal Energy, Entropy: **kJ/kg** (converted internally to J/kg)
  - Specific Heat: **kJ/kg/K** (converted internally to J/kg/K)

#### `PropsSI(output, name1, value1, name2, value2, fluid)`

- Computes thermodynamic properties of real fluids using **SI units** (no conversion).
- Standard CoolProp function.
- **Units used:**
  - Temperature: **Kelvin (K)**
  - Pressure: **Pascal (Pa)**
  - Enthalpy, Internal Energy, Entropy: **J/kg**
  - Specific Heat: **J/kg/K**

#### `Props1(output, fluid)` and `Props1SI(output, fluid)`

- Computes single-input properties (critical properties, molar mass, etc.)
- `Props1` returns values in **engineering units** (°C, bar)
- `Props1SI` returns values in **SI units** (K, Pa)

#### `Phase(name1, value1, name2, value2, fluid)` and `PhaseSI(...)`

- Determines the phase of a fluid at given conditions
- `Phase` accepts inputs in **engineering units** (°C, bar)
- `PhaseSI` accepts inputs in **SI units** (K, Pa)
- Returns phase string: "liquid", "gas", "supercritical", "twophase", etc.

### Example Usage:

```excel
=Props("H", "T", 25, "P", 1.01325, "Water")
```

This will return the enthalpy (H) of water at **25°C** and **1.01325 bar** in kJ/kg.

For SI units:

```excel
=PropsSI("H", "T", 298.15, "P", 101325, "Water")
```

This will return the enthalpy (H) of water at **298.15 K** and **101325 Pa** in J/kg.

Single-input properties:

```excel
=Props1("Tcrit", "Water")
```

Returns the critical temperature of water in **°C**.

Phase determination:

```excel
=Phase("T", 25, "P", 1.01325, "Water")
```

Returns "liquid" for water at 25°C and 1.01325 bar.

It can be used in a parametric way for convenience and efficient calculations. This allows for modular approaches, being able to reuse part of the calculations, drag the formulas, etc.

For example using:

```excel
=Props(E1,B1,B2,C1,C2,A2)
```

![Parametric usage](https://github.com/Danisaski/dafer-tmpr/blob/main/imgs/screenshot.png)

### Humid Air Functions

#### `HAProps(output, name1, value1, name2, value2, name3, value3)`

- Computes thermodynamic properties of humid air using **engineering units**.
- Standard CoolProp function with engineering units instead of SI.
- **Units used:**
  - Temperature: **Celsius (°C)** (converted internally to Kelvin (K))
  - Pressure: **bar** (converted internally to Pascal (Pa))
  - Enthalpy, Specific Heat, Entropy: **kJ/kg** (converted internally to J/kg)
  - Humidity Ratio: **kg_water/kg_dry_air**

#### `HAPropsSI(output, name1, value1, name2, value2, name3, value3)`

- Computes thermodynamic properties of humid air using **SI units** (no conversion).
- Standard CoolProp function.
- **Units used:**
  - Temperature: **Kelvin (K)**
  - Pressure: **Pascal (Pa)**
  - Enthalpy, Specific Heat, Entropy: **J/kg**
  - Humidity Ratio: **kg_water/kg_dry_air**

### Example Usage:

```excel
=HAProps("W", "T", 25, "P", 1.01325, "R", 0.5)
```

This will return the humidity ratio (W) for air at **25°C**, **1.01325 bar**, and **50% relative humidity**.

For SI units:

```excel
=HAPropsSI("W", "T", 298.15, "P", 101325, "R", 0.5)
```

This will return the humidity ratio (W) for air at **298.15 K**, **101325 Pa**, and **50% relative humidity**.

### Additional Functions

#### `GetGlobalParamString(param)`

- Retrieves global CoolProp parameters
- Example: `=GetGlobalParamString("version")` returns the CoolProp version

#### `GetFluidParamString(fluid, param)`

- Retrieves fluid-specific parameters
- Example: `=GetFluidParamString("Water", "formula")` returns "H2O"

### Diagnostic Function

#### `CPropDiag()`

- Diagnostic function to check CoolProp DLL loading paths.
- Returns information about where the add-in is searching for `CoolProp.dll` and whether it was found.
- Useful for troubleshooting installation issues.

### Function Summary

| **Function**           | **Description**                | **Units**                    |
| ---------------------- | ------------------------------ | ---------------------------- |
| `Props`                | Real fluid properties          | Engineering (°C, bar, kJ/kg) |
| `PropsSI`              | Real fluid properties          | SI (K, Pa, J/kg)             |
| `Props1`               | Single-input properties        | Engineering (°C, bar)        |
| `Props1SI`             | Single-input properties        | SI (K, Pa)                   |
| `Phase`                | Phase determination            | Engineering (°C, bar)        |
| `PhaseSI`              | Phase determination            | SI (K, Pa)                   |
| `HAProps`              | Humid air properties           | Engineering (°C, bar, kJ/kg) |
| `HAPropsSI`            | Humid air properties           | SI (K, Pa, J/kg)             |
| `GetGlobalParamString` | Get CoolProp global parameters | N/A                          |
| `GetFluidParamString`  | Get fluid-specific parameters  | N/A                          |
| `CPropDiag`            | Diagnostic tool                | N/A                          |

### Available Properties and Aliases

All parameter names are **case-insensitive**. You can type them in any combination of uppercase/lowercase.

| **Property** | **Aliases**              | **Description**                      | **Units in Add-in** | **SI/CoolProp Units** |
| ------------ | ------------------------ | ------------------------------------ | ------------------- | --------------------- |
| **T**        | t, temp, temperature     | Temperature                          | °C                  | K                     |
| **P**        | p, pres, pressure        | Pressure                             | bar                 | Pa                    |
| **H**        | h, enth, enthalpy, Hmass | Mass-specific enthalpy               | kJ/kg               | J/kg                  |
| **S**        | s, entr, entropy, Smass  | Mass-specific entropy                | kJ/kg/K             | J/kg/K                |
| **U**        | u, internalenergy, Umass | Mass-specific internal energy        | kJ/kg               | J/kg                  |
| **D**        | rho, dens, Dmass         | Mass density                         | kg/m³               | kg/m³                 |
| **Q**        | q, quality, x            | Vapor quality (0=liquid, 1=vapor)    | -                   | -                     |
| **Cpmass**   | cp, cpmass               | Mass-specific constant pressure heat | kJ/kg/K             | J/kg/K                |
| **Cvmass**   | cv, cvmass               | Mass-specific constant volume heat   | kJ/kg/K             | J/kg/K                |
| **A**        | speed_of_sound           | Speed of sound                       | m/s                 | m/s                   |
| **MU**       | mu, viscosity            | Dynamic viscosity                    | Pa·s                | Pa·s                  |
| **K**        | k, conductivity          | Thermal conductivity                 | W/m/K               | W/m/K                 |
| **Tcrit**    | pcrit, T_critical        | Critical temperature                 | °C                  | K                     |
| **Pcrit**    | p_critical               | Critical pressure                    | bar                 | Pa                    |
| **MM**       | molar_mass               | Molar mass                           | kg/mol              | kg/mol                |

### Humid Air Properties

| **Property** | **Aliases**    | **Description**                      | **Units in Add-in** | **SI/CoolProp Units** |
| ------------ | -------------- | ------------------------------------ | ------------------- | --------------------- |
| **T**        | t, Tdb, T_db   | Dry bulb temperature                 | °C                  | K                     |
| **Twb**      | T_wb, wetbulb  | Wet bulb temperature                 | °C                  | K                     |
| **Tdp**      | dewpoint, T_dp | Dew point temperature                | °C                  | K                     |
| **P**        | p              | Atmospheric pressure                 | bar                 | Pa                    |
| **P_w**      |                | Water vapor partial pressure         | bar                 | Pa                    |
| **R**        | rh, RH, relhum | Relative humidity (0-1)              | -                   | -                     |
| **W**        | omega, humrat  | Humidity ratio (kg water/kg dry air) | kg/kg               | kg/kg                 |
| **Hha**      |                | Humid air enthalpy per kg dry air    | kJ/kg               | J/kg                  |
| **Hda**      |                | Dry air enthalpy per kg dry air      | kJ/kg               | J/kg                  |
| **Vha**      |                | Humid air specific volume            | m³/kg               | m³/kg                 |
| **Dha**      | rhoha          | Humid air density                    | kg/m³               | kg/m³                 |
| **Cha**      | Cpha           | Humid air specific heat              | kJ/kg/K             | J/kg/K                |

### Available Fluids

#### Pure Fluids

1-Butene, Acetone, Air, Ammonia, Argon, Benzene, CarbonDioxide, CarbonMonoxide, CarbonylSulfide, cis-2-Butene, CycloHexane, Cyclopentane, CycloPropane, D4, D5, D6, Deuterium, Dichloroethane, DiethylEther, DimethylCarbonate, DimethylEther, Ethane, Ethanol, EthylBenzene, Ethylene, EthyleneOxide, Fluorine, HeavyWater, Helium, HFE143m, Hydrogen, HydrogenChloride, HydrogenSulfide, IsoButane, IsoButene, Isohexane, Isopentane, Krypton, m-Xylene, MD2M, MD3M, MD4M, MDM, Methane, Methanol, MethylLinoleate, MethylLinolenate, MethylOleate, MethylPalmitate, MethylStearate, MM, n-Butane, n-Decane, n-Dodecane, n-Heptane, n-Hexane, n-Nonane, n-Octane, n-Pentane, n-Propane, n-Undecane, Neon, Neopentane, Nitrogen, NitrousOxide, Novec649, o-Xylene, OrthoDeuterium, OrthoHydrogen, Oxygen, p-Xylene, ParaDeuterium, ParaHydrogen, Propylene, Propyne, R11, R113, R114, R115, R116, R12, R123, R1233zd(E), R1234yf, R1234ze(E), R1234ze(Z), R124, R125, R13, R134a, R13I1, R14, R141b, R142b, R143a, R152A, R161, R21, R218, R22, R227EA, R23, R236EA, R236FA, R245ca, R245fa, R32, R365MFC, R40, R404A, R407C, R41, R410A, R507A, RC318, SES36, SulfurDioxide, SulfurHexafluoride, Toluene, trans-2-Butene, Water, Xenon

#### Predefined Mixtures

Air.mix, Amarillo.mix, Ekofisk.mix, GulfCoast.mix, GulfCoastGas(NIST1).mix, HighCO2.mix, HighN2.mix, NaturalGasSample.mix, R401A.mix, R401B.mix, R401C.mix, R402A.mix, R402B.mix, R403A.mix, R403B.mix, R404A.mix, R405A.mix, R406A.mix, R407A.mix, R407B.mix, R407C.mix, R407D.mix, R407E.mix, R407F.mix, R408A.mix, R409A.mix, R409B.mix, R410A.mix, R410B.mix, R411A.mix, R411B.mix, R412A.mix, R413A.mix, R414A.mix, R414B.mix, R415A.mix, R415B.mix, R416A.mix, R417A.mix, R417B.mix, R417C.mix, R418A.mix, R419A.mix, R419B.mix, R420A.mix, R421A.mix, R421B.mix, R422A.mix, R422B.mix, R422C.mix, R422D.mix, R422E.mix, R423A.mix, R424A.mix, R425A.mix, R426A.mix, R427A.mix, R428A.mix, R429A.mix, R430A.mix, R431A.mix, R432A.mix, R433A.mix, R433B.mix, R433C.mix, R434A.mix, R435A.mix, R436A.mix, R436B.mix, R437A.mix, R438A.mix, R439A.mix, R440A.mix, R441A.mix, R442A.mix, R443A.mix, R444A.mix, R444B.mix, R445A.mix, R446A.mix, R447A.mix, R448A.mix, R449A.mix, R449B.mix, R450A.mix, R451A.mix, R451B.mix, R452A.mix, R453A.mix, R454A.mix, R454B.mix, R500.mix, R501.mix, R502.mix, R503.mix, R504.mix, R507A.mix, R508A.mix, R508B.mix, R509A.mix, R510A.mix, R511A.mix, R512A.mix, R513A.mix, TypicalNaturalGas.mix

You can also use CoolProp's mixture notation for custom mixtures, e.g., `"R32[0.5]&R125[0.5]"` for a 50/50 mix.

### Complete Parameter List

The following parameters can be used in the `output`, `name1`, `name2`, and `name3` arguments. Parameters are **case-insensitive**:

**State Properties:**
A, ACENTRIC, acentric, ALPHA0, alpha0, ALPHAR, alphar, BVIRIAL, Bvirial, C, CONDUCTIVITY, conductivity, CP0MASS, Cp0mass, CP0MOLAR, Cp0molar, CPMASS, Cpmass, CPMOLAR, Cpmolar, CVIRIAL, Cvirial, CVMASS, Cvmass, CVMOLAR, Cvmolar, D, DELTA, Delta, DIPOLE_MOMENT, dipole_moment, DMASS, Dmass, DMOLAR, Dmolar

**Derivatives:**
DALPHA0_DDELTA_CONSTTAU, dalpha0_ddelta_consttau, DALPHA0_DTAU_CONSTDELTA, dalpha0_dtau_constdelta, DALPHAR_DDELTA_CONSTTAU, dalphar_ddelta_consttau, DALPHAR_DTAU_CONSTDELTA, dalphar_dtau_constdelta, DBVIRIAL_DT, dBvirial_dT, DCVIRIAL_DT, dCvirial_dT

**Energy & Entropy:**
FH, FUNDAMENTAL_DERIVATIVE_OF_GAS_DYNAMICS, fundamental_derivative_of_gas_dynamics, G, GAS_CONSTANT, gas_constant, GMASS, Gmass, GMOLAR, Gmolar, H, HH, HMASS, Hmass, HMOLAR, Hmolar, S, SMASS, Smass, SMOLAR, Smolar, SMOLAR_RESIDUAL, Smolar_residual, U, UMASS, Umass, UMOLAR, Umolar

**Environmental:**
GWP20, GWP100, GWP500, ODP

**Transport & Dimensionless:**
I, ISOBARIC_EXPANSION_COEFFICIENT, isobaric_expansion_coefficient, ISOTHERMAL_COMPRESSIBILITY, isothermal_compressibility, L, PRANDTL, Prandtl, VISCOSITY, viscosity, Z

**Critical, Triple & Reducing:**
PCRIT, Pcrit, P_CRITICAL, p_critical, PMAX, P_MAX, P_max, pmax, PMIN, P_MIN, P_min, pmin, PTRIPLE, P_TRIPLE, p_triple, ptriple, P_REDUCING, p_reducing, TCRIT, Tcrit, T_CRITICAL, T_critical, TMAX, T_MAX, T_max, Tmax, TMIN, T_MIN, T_min, Tmin, TTRIPLE, T_TRIPLE, t_triple, Ttriple, T_FREEZE, T_freeze, t_freeze, T_REDUCING, T_reducing, t_reducing, RHOCRIT, rhocrit, RHOMASS_CRITICAL, rhomass_critical, RHOMASS_REDUCING, rhomass_reducing, RHOMOLAR_CRITICAL, rhomolar_critical, RHOMOLAR_REDUCING, rhomolar_reducing

**Molecular & Mixture:**
M, MOLARMASS, molarmass, MOLAR_MASS, molar_mass, MOLEMASS, molemass, FRACTION_MAX, fraction_max, FRACTION_MIN, fraction_min

**Phase:**
O, PHASE, Phase, PH, PIP, Q

**State Variables:**
P, T, TAU, Tau, V

**Surface:**
SURFACE_TENSION, surface_tension, SPEED_OF_SOUND, speed_of_sound

For more details, visit the [CoolProp documentation](http://www.coolprop.org/).

## If building from source:

### Quick Build (Windows)

1. Make sure you have .NET SDK installed
2. Run `build.bat` in the project directory
3. The compiled files will be in `compiled\net48\` directory

### Manual Build

```bash
cd src
dotnet restore
dotnet build -c Release -p:Platform=x64
```

## License

This project is provided as-is. CoolProp is licensed under the MIT License. See the CoolProp project for details.

## Credits

- [CoolProp](http://www.coolprop.org/) - Open-source thermodynamic and transport properties library
- [Excel-DNA](https://excel-dna.net/) - .NET library for building Excel add-ins
