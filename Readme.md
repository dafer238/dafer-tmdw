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
- **Supports Excel range inputs** - any numeric parameter can be a cell range, and the function will automatically calculate results for all values and spill them as an array.
- **Units used:**
  - Temperature: **Celsius (°C)** (converted internally to Kelvin (K))
  - Pressure: **bar** (converted internally to Pascal (Pa))
  - Enthalpy, Internal Energy, Entropy: **kJ/kg** (converted internally to J/kg)
  - Specific Heat: **kJ/kg/K** (converted internally to J/kg/K)

#### `PropsSI(output, name1, value1, name2, value2, fluid)`

- Computes thermodynamic properties of real fluids using **SI units** (no conversion).
- Standard CoolProp function.
- **Supports Excel range inputs** - any numeric parameter can be a cell range, and the function will automatically calculate results for all values and spill them as an array.
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

**Array calculations** - compute properties for multiple values at once:

```excel
=Props("H", "T", 25, "P", A1:A10, "Water")
```

This will return an array of 10 enthalpy values, one for each pressure in the range A1:A10, automatically spilling the results to adjacent cells below (column orientation).

**Row arrays** - when the input is a row range, the output will also spill horizontally:

```excel
=Props("H", "T", 25, "P", A1:J1, "Water")
```

This will return a row of 10 enthalpy values, spilling horizontally across columns.

**Multiple array inputs** - when using multiple ranges, they must have the same length:

```excel
=Props("H", "T", B1:B10, "P", A1:A10, "Water")
```

This will compute 10 enthalpy values using paired temperatures from B1:B10 and pressures from A1:A10.

**Mixed scalar and array inputs**:

```excel
=Props("H", "P", 1.01325, "T", A1:A5, "Water")
```

The scalar pressure value (1.01325 bar) will be used for all 5 temperature values in A1:A5.

Single-input properties:

```excel
=Props1("Tcrit", "Water")
```

Returns the critical temperature of water in **°C** (approximately 373.95°C).

```excel
=Props1("Tcrit", "R410A")
```

Returns the critical temperature of R410A in **°C** (approximately 71.35°C, which is 344.5 K in SI units).

```excel
=Props1("Pcrit", "Water")
```

Returns the critical pressure of water in **bar** (approximately 220.6 bar).

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

![Parametric usage](https://github.com/dafer238/dafer-tmdw/blob/main/imgs/screenshot.png)

### Humid Air Functions

#### `HAProps(output, name1, value1, name2, value2, name3, value3)`

- Computes thermodynamic properties of humid air using **engineering units**.
- Standard CoolProp function with engineering units instead of SI.
- **Supports Excel range inputs** - any numeric parameter can be a cell range, and the function will automatically calculate results for all values and spill them as an array.
- **Units used:**
  - Temperature: **Celsius (°C)** (converted internally to Kelvin (K))
  - Pressure: **bar** (converted internally to Pascal (Pa))
  - Enthalpy, Specific Heat, Entropy: **kJ/kg** (converted internally to J/kg)
  - Humidity Ratio: **kg_water/kg_dry_air**

#### `HAPropsSI(output, name1, value1, name2, value2, name3, value3)`

- Computes thermodynamic properties of humid air using **SI units** (no conversion).
- Standard CoolProp function.
- **Supports Excel range inputs** - any numeric parameter can be a cell range, and the function will automatically calculate results for all values and spill them as an array.
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

**Array calculations** - compute humid air properties for multiple values:

```excel
=HAProps("W", "T", A1:A10, "P", 1.01325, "R", 0.5)
```

This will return an array of 10 humidity ratio values, one for each temperature in A1:A10, spilling vertically.

**Row arrays** - when the input is a row range, the output will also spill horizontally:

```excel
=HAProps("W", "T", A1:J1, "P", 1.01325, "R", 0.5)
```

This will return a row of 10 humidity ratio values, spilling horizontally across columns.

**Multiple array inputs** - all ranges must have the same length:

```excel
=HAProps("Hda", "T", A1:A5, "P", 1.01325, "R", B1:B5)
```

This will compute 5 enthalpy values using paired temperatures from A1:A5 and relative humidities from B1:B5.

### Additional Functions

#### `GetGlobalParam(param)`

- Retrieves global CoolProp parameters
- Example: `=GetGlobalParam("version")` returns the CoolProp version
- Example: `=GetGlobalParam("gitrevision")` returns the Git revision

#### `GetFluidParam(fluid, param)`

- Retrieves fluid-specific parameters
- Example: `=GetFluidParam("Water", "formula")` returns "H2O"
- Example: `=GetFluidParam("Water", "CAS")` returns CAS registry number

#### `MixtureString(elements, fractions)`

- Creates a mixture string for CoolProp from element names and their mole/mass fractions
- Example: `=MixtureString(A1:A2, B1:B2)` where A1:A2 contains ["R32", "R125"] and B1:B2 contains [0.5, 0.5]
- Returns: "HEOS::R32[0.5]&R125[0.5]"

### Diagnostic Function

#### `CPropDiag()`

- Diagnostic function to check CoolProp DLL loading paths.
- Returns information about where the add-in is searching for `CoolProp.dll` and whether it was found.
- Useful for troubleshooting installation issues.

### Function Summary

| **Function**     | **Description**                | **Units**                    |
| ---------------- | ------------------------------ | ---------------------------- |
| `Props`          | Real fluid properties          | Engineering (°C, bar, kJ/kg) |
| `PropsSI`        | Real fluid properties          | SI (K, Pa, J/kg)             |
| `Props1`         | Single-input properties        | Engineering (°C, bar)        |
| `Props1SI`       | Single-input properties        | SI (K, Pa)                   |
| `Phase`          | Phase determination            | Engineering (°C, bar)        |
| `PhaseSI`        | Phase determination            | SI (K, Pa)                   |
| `HAProps`        | Humid air properties           | Engineering (°C, bar, kJ/kg) |
| `HAPropsSI`      | Humid air properties           | SI (K, Pa, J/kg)             |
| `GetGlobalParam` | Get CoolProp global parameters | N/A                          |
| `GetFluidParam`  | Get fluid-specific parameters  | N/A                          |
| `MixtureString`  | Create mixture strings         | N/A                          |
| `CPropDiag`      | Diagnostic tool                | N/A                          |

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
<details>
  <summary><h3>Unfold for a complete list of accepted inputs</h3></summary>
  
### Complete Parameter Reference

#### Props/PropsSI/Props1/Props1SI Parameters

All parameters are **case-insensitive**. The following table shows all recognized parameter variations:

| **Output Name**                            | **All Accepted Aliases (case-insensitive)**                                     | **Description**                               | **Units (Props)** | **Units (PropsSI)** |
| ------------------------------------------ | ------------------------------------------------------------------------------- | --------------------------------------------- | ----------------- | ------------------- |
| **T**                                      | T, TEMP, TEMPERATURE                                                            | Temperature                                   | °C                | K                   |
| **P**                                      | P, PRES, PRESSURE                                                               | Pressure                                      | bar               | Pa                  |
| **D**                                      | D, RHO, DENS, DENSITY, DMASS, D_MASS, RHOMASS                                   | Mass density                                  | kg/m³             | kg/m³               |
| **Dmolar**                                 | DMOLAR, RHOMOLAR, DMOL, D_MOL, D_MOLAR                                          | Molar density                                 | mol/m³            | mol/m³              |
| **H**                                      | H, ENTH, ENTHALPY, HMASS, H_MASS                                                | Mass specific enthalpy                        | kJ/kg             | J/kg                |
| **Hmolar**                                 | HMOLAR, H_MOLAR                                                                 | Molar specific enthalpy                       | kJ/mol            | J/mol               |
| **S**                                      | S, ENTR, ENTROPY, SMASS, S_MASS                                                 | Mass specific entropy                         | kJ/(kg·K)         | J/(kg·K)            |
| **Smolar**                                 | SMOLAR, S_MOLAR                                                                 | Molar specific entropy                        | kJ/(mol·K)        | J/(mol·K)           |
| **Smolar_residual**                        | SMOLAR_RESIDUAL, SMOLARRESIDUAL                                                 | Residual molar entropy                        | kJ/(mol·K)        | J/(mol·K)           |
| **U**                                      | U, INTERNALENERGY, UMASS, U_MASS                                                | Mass specific internal energy                 | kJ/kg             | J/kg                |
| **Umolar**                                 | UMOLAR, U_MOLAR                                                                 | Molar specific internal energy                | kJ/mol            | J/mol               |
| **Q**                                      | Q, QUALITY, X, VAPORFRACTION, VAPOR_FRACTION                                    | Vapor quality (0=liquid, 1=vapor)             | -                 | -                   |
| **Cpmass**                                 | CPMASS, CP, C                                                                   | Mass specific heat (constant P)               | kJ/(kg·K)         | J/(kg·K)            |
| **Cvmass**                                 | CVMASS, CV                                                                      | Mass specific heat (constant V)               | kJ/(kg·K)         | J/(kg·K)            |
| **Cpmolar**                                | CPMOLAR, CPMOL                                                                  | Molar specific heat (constant P)              | kJ/(mol·K)        | J/(mol·K)           |
| **Cvmolar**                                | CVMOLAR, CVMOL                                                                  | Molar specific heat (constant V)              | kJ/(mol·K)        | J/(mol·K)           |
| **Cp0mass**                                | CP0MASS                                                                         | Ideal gas mass specific heat                  | kJ/(kg·K)         | J/(kg·K)            |
| **Cp0molar**                               | CP0MOLAR                                                                        | Ideal gas molar specific heat                 | kJ/(mol·K)        | J/(mol·K)           |
| **G**                                      | G, GMASS, GIBBS                                                                 | Mass specific Gibbs energy                    | kJ/kg             | J/kg                |
| **Gmolar**                                 | GMOLAR                                                                          | Molar specific Gibbs energy                   | kJ/mol            | J/mol               |
| **A**                                      | A, SPEED_OF_SOUND, SPEEDOFSOUND, W                                              | Speed of sound                                | m/s               | m/s                 |
| **viscosity**                              | MU, VISCOSITY, V                                                                | Dynamic viscosity                             | Pa·s              | Pa·s                |
| **conductivity**                           | K, CONDUCTIVITY, L                                                              | Thermal conductivity                          | W/(m·K)           | W/(m·K)             |
| **surface_tension**                        | SURFACE_TENSION, SURFACETENSION, SIGMA, I                                       | Surface tension                               | N/m               | N/m                 |
| **Z**                                      | Z, COMPRESSIBILITY, COMPRESSIBILITYFACTOR                                       | Compressibility factor                        | -                 | -                   |
| **Prandtl**                                | PRANDTL                                                                         | Prandtl number                                | -                 | -                   |
| **Tau**                                    | TAU                                                                             | Reduced reciprocal temperature (Tc/T)         | -                 | -                   |
| **Delta**                                  | DELTA                                                                           | Reduced density (ρ/ρc)                        | -                 | -                   |
| **Alpha0**                                 | ALPHA0                                                                          | Ideal Helmholtz energy                        | -                 | -                   |
| **Alphar**                                 | ALPHAR                                                                          | Residual Helmholtz energy                     | -                 | -                   |
| **HELMHOLTZMASS**                          | HELMHOLTZMASS, HH                                                               | Mass specific Helmholtz energy                | J/kg              | J/kg                |
| **HELMHOLTZMOLAR**                         | HELMHOLTZMOLAR                                                                  | Molar specific Helmholtz energy               | J/mol             | J/mol               |
| **Bvirial**                                | BVIRIAL                                                                         | Second virial coefficient                     | -                 | -                   |
| **Cvirial**                                | CVIRIAL                                                                         | Third virial coefficient                      | -                 | -                   |
| **DBVIRIAL_DT**                            | DBVIRIAL_DT, DBVIRIAL, DBVIRIALDT                                               | Derivative of 2nd virial coeff w.r.t. T       | -                 | -                   |
| **DCVIRIAL_DT**                            | DCVIRIAL_DT, DCVIRIAL, DCVIRIALDT                                               | Derivative of 3rd virial coeff w.r.t. T       | -                 | -                   |
| **isentropic_expansion_coefficient**       | GAMMA, ISENTROPICEXPANSIONCOEFFICIENT                                           | Isentropic expansion coefficient              | -                 | -                   |
| **isobaric_expansion_coefficient**         | ISOBARIC_EXPANSION_COEFFICIENT, ISOBARICEXPANSION, ISOBARICEXPANSIONCOEFFICIENT | Isobaric expansion coefficient                | 1/K               | 1/K                 |
| **isothermal_compressibility**             | ISOTHERMAL_COMPRESSIBILITY, ISOTHERMALCOMPRESSIBILITY                           | Isothermal compressibility                    | 1/Pa              | 1/Pa                |
| **fundamental_derivative_of_gas_dynamics** | FUNDAMENTAL_DERIVATIVE_OF_GAS_DYNAMICS, FUNDAMENTALDERIVATIVE, FH               | Fundamental derivative of gas dynamics        | -                 | -                   |
| **acentric**                               | ACENTRIC, ACENTRIC_FACTOR, ACENTRICFACTOR, OMEGA                                | Acentric factor                               | -                 | -                   |
| **DIPOLE_MOMENT**                          | DIPOLE_MOMENT, DIPOLEMOMENT, DIPOLE                                             | Dipole moment                                 | C·m               | C·m                 |
| **gas_constant**                           | GAS_CONSTANT, GASCONSTANT                                                       | Molar gas constant                            | J/(mol·K)         | J/(mol·K)           |
| **M**                                      | MM, MOLAR_MASS, MOLARMASS, MOLEMASS                                             | Molar mass                                    | kg/mol            | kg/mol              |
| **Tcrit**                                  | TCRIT, T_CRITICAL, TCRITICAL                                                    | Critical temperature                          | °C                | K                   |
| **Pcrit**                                  | PCRIT, P_CRITICAL, PCRITICAL                                                    | Critical pressure                             | bar               | Pa                  |
| **rhocrit**                                | RHOCRIT, RHOCRITICAL, DCRIT                                                     | Mass density at critical point                | kg/m³             | kg/m³               |
| **rhomass_critical**                       | RHOMASS_CRITICAL, RHOMASSCRITICAL                                               | Mass density at critical point                | kg/m³             | kg/m³               |
| **rhomolar_critical**                      | RHOMOLAR_CRITICAL, RHOMOLARCRITICAL                                             | Molar density at critical point               | mol/m³            | mol/m³              |
| **Tmax**                                   | TMAX, T_MAX                                                                     | Maximum temperature limit                     | °C                | K                   |
| **Tmin**                                   | TMIN, T_MIN                                                                     | Minimum temperature limit                     | °C                | K                   |
| **pmax**                                   | PMAX, P_MAX                                                                     | Maximum pressure limit                        | bar               | Pa                  |
| **pmin**                                   | PMIN, P_MIN                                                                     | Minimum pressure limit                        | bar               | Pa                  |
| **Ttriple**                                | TTRIPLE, T_TRIPLE, TTRIP                                                        | Triple point temperature                      | °C                | K                   |
| **ptriple**                                | PTRIPLE, P_TRIPLE, PTRIP                                                        | Triple point pressure                         | bar               | Pa                  |
| **T_freeze**                               | T_FREEZE, TFREEZE, FREEZING_TEMPERATURE                                         | Freezing temperature                          | °C                | K                   |
| **T_reducing**                             | T_REDUCING, TREDUCING                                                           | Reducing point temperature                    | °C                | K                   |
| **p_reducing**                             | P_REDUCING, PREDUCING                                                           | Reducing point pressure                       | bar               | Pa                  |
| **rhomass_reducing**                       | RHOMASS_REDUCING, RHOMASSREDUCING                                               | Mass density at reducing point                | kg/m³             | kg/m³               |
| **rhomolar_reducing**                      | RHOMOLAR_REDUCING, RHOMOLARREDUCING                                             | Molar density at reducing point               | mol/m³            | mol/m³              |
| **fraction_max**                           | FRACTION_MAX, FRACTIONMAX                                                       | Maximum fraction for incompressible solutions | -                 | -                   |
| **fraction_min**                           | FRACTION_MIN, FRACTIONMIN                                                       | Minimum fraction for incompressible solutions | -                 | -                   |
| **GWP20**                                  | GWP20, GWP_20                                                                   | 20-year global warming potential              | -                 | -                   |
| **GWP100**                                 | GWP100, GWP_100                                                                 | 100-year global warming potential             | -                 | -                   |
| **GWP500**                                 | GWP500, GWP_500                                                                 | 500-year global warming potential             | -                 | -                   |
| **ODP**                                    | ODP, OZONEDEPLETIONPOTENTIAL                                                    | Ozone depletion potential                     | -                 | -                   |
| **Phase**                                  | PHASE                                                                           | Phase identification                          | -                 | -                   |
| **PIP**                                    | PIP                                                                             | Phase identification parameter                | -                 | -                   |

**Helmholtz Energy Derivatives** (advanced users):
- DALPHA0_DDELTA_CONSTTAU, DALPHA0_DDELTA
- DALPHA0_DTAU_CONSTDELTA, DALPHA0_DTAU  
- DALPHAR_DDELTA_CONSTTAU, DALPHAR_DDELTA
- DALPHAR_DTAU_CONSTDELTA, DALPHAR_DTAU

#### HAProps/HAPropsSI Parameters

All parameters for humid air calculations are **case-insensitive**:

| **Output Name** | **All Accepted Aliases (case-insensitive)**                               | **Description**                    | **Units (HAProps)**     | **Units (HAPropsSI)**   |
| --------------- | ------------------------------------------------------------------------- | ---------------------------------- | ----------------------- | ----------------------- |
| **T**           | T, TDB, T_DB, TEMP, TEMPERATURE, DRYBULB, DRYBULBTEMP, DRYBULBTEMPERATURE | Dry-bulb temperature               | °C                      | K                       |
| **Twb**         | B, TWB, T_WB, WETBULB, WETBULBTEMP, WETBULBTEMPERATURE                    | Wet-bulb temperature               | °C                      | K                       |
| **Tdp**         | D, TDP, T_DP, DEWPOINT, DEWPOINTTEMP, DEWPOINTTEMPERATURE                 | Dew-point temperature              | °C                      | K                       |
| **P**           | P, PRESSURE, PRES                                                         | Atmospheric pressure               | bar                     | Pa                      |
| **P_w**         | P_W, PW, PARTIALPRESSURE, WATERPRESSURE                                   | Water vapor partial pressure       | bar                     | Pa                      |
| **R**           | R, RH, RELHUM, RELATIVEHUMIDITY                                           | Relative humidity (0-1)            | -                       | -                       |
| **W**           | W, OMEGA, HUMRAT, HUMIDITYRATIO, MIXINGRATIO                              | Humidity ratio                     | kg water/kg dry air     | kg water/kg dry air     |
| **Psi_w**       | PSI_W, PSIW, Y                                                            | Water mole fraction                | mol water/mol humid air | mol water/mol humid air |
| **Hda**         | H, HDA, H_DA, ENTHALPY                                                    | Enthalpy per unit dry air          | kJ/kg dry air           | J/kg dry air            |
| **Hha**         | HHA, H_HA                                                                 | Enthalpy per unit humid air        | kJ/kg humid air         | J/kg humid air          |
| **Sda**         | S, SDA, S_DA, ENTROPY                                                     | Entropy per unit dry air           | kJ/(kg·K) dry air       | J/(kg·K) dry air        |
| **Sha**         | SHA, S_HA                                                                 | Entropy per unit humid air         | kJ/(kg·K) humid air     | J/(kg·K) humid air      |
| **Vda**         | V, VDA, V_DA                                                              | Volume per unit dry air            | m³/kg dry air           | m³/kg dry air           |
| **Vha**         | VHA, V_HA                                                                 | Volume per unit humid air          | m³/kg humid air         | m³/kg humid air         |
| **Cda**         | C, CP, CDA, CPDA, CP_DA                                                   | Specific heat per unit dry air     | kJ/(kg·K) dry air       | J/(kg·K) dry air        |
| **Cha**         | CHA, CPHA, CP_HA                                                          | Specific heat per unit humid air   | kJ/(kg·K) humid air     | J/(kg·K) humid air      |
| **CV**          | CV, CVMASS                                                                | Constant volume heat per dry air   | kJ/(kg·K) dry air       | J/(kg·K) dry air        |
| **CVha**        | CVHA, CV_HA                                                               | Constant volume heat per humid air | kJ/(kg·K) humid air     | J/(kg·K) humid air      |
| **K**           | K, CONDUCTIVITY, THERMALCONDUCTIVITY                                      | Thermal conductivity               | W/(m·K)                 | W/(m·K)                 |
| **MU**          | M, MU, VISC, VISCOSITY, DYNAMICVISCOSITY                                  | Dynamic viscosity                  | Pa·s                    | Pa·s                    |
| **Z**           | Z, COMPRESSIBILITY, COMPRESSIBILITYFACTOR                                 | Compressibility factor             | -                       | -                       |

</details>

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

CoolProp is licensed under the MIT License. See the CoolProp project for details. This project is licensed under GNU GPL v3 license.

## Credits

- [CoolProp](http://www.coolprop.org/) - Open-source thermodynamic and transport properties library
- [Excel-DNA](https://excel-dna.net/) - .NET library for building Excel add-ins
