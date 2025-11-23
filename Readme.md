# A CoolProp wrapper for Excel in engineering units

## Installation

[Latest release](https://github.com/Danisaski/dafer-tmpr/releases) is available for a portable standalone version. If you are **not sure which .NET distribution is installed in your machine**, if any at all, **[download](https://github.com/Danisaski/dafer-tmpr/releases/tag/v0.1.3) the portable rolling release version**. If you have .NET SDK installed you could check which version to download via:

```bash
dotnet --list-runtimes
```

After downloading, extract the zip file with the contents; `Coolprop.dll`, `CoolpropWrapper.dll` and `CoolpropWrapper.xll` (and `CoolPropWrapper.dna` for the portable version), place the three (four) files **together** in the directory of your choice. Finally import in Excel the `.xll` file via: File (Archivo) -> Options (Opciones) -> Addins (Complementos) -> Import (Ir...) -> Browse... (Examinar...) and select the `.xll` file.

## Usage

CoolPropWrapper allows you to compute thermodynamic properties of real fluids and humid air using CoolProp in Excel. The functions are available in two variants:

- **Engineering Units** (default): Temperature in °C, Pressure in bar, Energy in kJ/kg, etc.
- **SI Units**: Temperature in K, Pressure in Pa, Energy in J/kg, etc.

### Real Fluid Functions

#### `CProp(output, name1, value1, name2, value2, fluid)` or `CProp_E(...)`

- Computes thermodynamic properties of real fluids using **engineering units**.
- **Units used:**
  - Temperature: **Celsius (°C)** (converted internally to Kelvin (K))
  - Pressure: **bar** (converted internally to Pascal (Pa))
  - Enthalpy, Internal Energy, Entropy: **kJ/kg** (converted internally to J/kg)
  - Specific Heat: **kJ/kg/K** (converted internally to J/kg/K)

#### `CProp_SI(output, name1, value1, name2, value2, fluid)`

- Computes thermodynamic properties of real fluids using **SI units** (no conversion).
- **Units used:**
  - Temperature: **Kelvin (K)**
  - Pressure: **Pascal (Pa)**
  - Enthalpy, Internal Energy, Entropy: **J/kg**
  - Specific Heat: **J/kg/K**

### Example Usage:

```excel
=CProp("H", "T", 25, "P", 1.01325, "Water")
```

This will return the enthalpy (H) of water at **25°C** and **1.01325 bar**.

For SI units:

```excel
=CProp_SI("H", "T", 298.15, "P", 101325, "Water")
```

This will return the enthalpy (H) of water at **298.15 K** and **101325 Pa** in J/kg.

It can be used in a parametric way for convenience and efficient calculations. This allows for modular approaches, being able to reuse part of the calculations, drag the formulas, etc.

For example using:

```excel
=CProp(E1,B1,B2,C1,C2,A2)
```

![Parametric usage](https://github.com/Danisaski/dafer-tmpr/blob/main/imgs/screenshot.png)

### Humid Air Functions

#### `CPropHA(output, name1, value1, name2, value2, name3, value3)` or `CPropHA_E(...)`

- Computes thermodynamic properties of humid air using **engineering units**.
- **Units used:**
  - Temperature: **Celsius (°C)** (converted internally to Kelvin (K))
  - Pressure: **bar** (converted internally to Pascal (Pa))
  - Enthalpy, Specific Heat, Entropy: **kJ/kg** (converted internally to J/kg)
  - Humidity Ratio: **kg_water/kg_dry_air**

#### `CPropHA_SI(output, name1, value1, name2, value2, name3, value3)`

- Computes thermodynamic properties of humid air using **SI units** (no conversion).
- **Units used:**
  - Temperature: **Kelvin (K)**
  - Pressure: **Pascal (Pa)**
  - Enthalpy, Specific Heat, Entropy: **J/kg**
  - Humidity Ratio: **kg_water/kg_dry_air**

### Example Usage:

```excel
=CPropHA("W", "T", 25, "P", 1.01325, "RH", 0.5)
```

This will return the humidity ratio (W) for air at **25°C**, **1.01325 bar**, and **50% relative humidity**.

For SI units:

```excel
=CPropHA_SI("W", "T", 298.15, "P", 101325, "R", 0.5)
```

This will return the humidity ratio (W) for air at **298.15 K**, **101325 Pa**, and **50% relative humidity**.

### Diagnostic Function

#### `CPropDiag()`

- Diagnostic function to check CoolProp DLL loading paths.
- Returns information about where the add-in is searching for `CoolProp.dll` and whether it was found.
- Useful for troubleshooting installation issues.

### Function Summary

| **Function** | **Description**                 | **Units**                    |
| ------------ | ------------------------------- | ---------------------------- |
| `CProp`      | Real fluid properties (default) | Engineering (°C, bar, kJ/kg) |
| `CProp_E`    | Real fluid properties           | Engineering (°C, bar, kJ/kg) |
| `CProp_SI`   | Real fluid properties           | SI (K, Pa, J/kg)             |
| `CPropHA`    | Humid air properties (default)  | Engineering (°C, bar, kJ/kg) |
| `CPropHA_E`  | Humid air properties            | Engineering (°C, bar, kJ/kg) |
| `CPropHA_SI` | Humid air properties            | SI (K, Pa, J/kg)             |
| `CPropDiag`  | Diagnostic tool                 | N/A                          |

### Available Properties and Aliases

| **Property**           | **Aliases**              | **Description**                        | **Units in Add-in** | **SI/CoolProp Units** |
| ---------------------- | ------------------------ | -------------------------------------- | ------------------- | --------------------- |
| Temperature            | `T`, `temperature`       | Absolute temperature                   | °C                  | K                     |
| Pressure               | `P`, `pressure`          | Absolute pressure                      | bar                 | Pa                    |
| Density                | `D`, `rho`, `dmass`      | Mass density                           | kg/m³               | kg/m³                 |
| Molar Density          | `Dmolar`                 | Molar density                          | mol/m³              | mol/m³                |
| Enthalpy               | `H`, `hmass`, `enthalpy` | Specific enthalpy                      | kJ/kg               | J/kg                  |
| Internal Energy        | `U`, `umass`             | Specific internal energy               | kJ/kg               | J/kg                  |
| Entropy                | `S`, `smass`             | Specific entropy                       | kJ/kg·K             | J/kg·K                |
| Specific Heat (Cp)     | `Cpmass`, `Cp`           | Heat capacity at constant pressure     | kJ/kg·K             | J/kg·K                |
| Specific Heat (Cv)     | `Cvmass`, `Cv`           | Heat capacity at constant volume       | kJ/kg·K             | J/kg·K                |
| Quality                | `Q`                      | Vapor quality (mass fraction of vapor) | (unitless)          | (unitless)            |
| Speed of Sound         | `A`, `speed_of_sound`    | Speed of sound in the fluid            | m/s                 | m/s                   |
| Thermal Conductivity   | `K`, `conductivity`      | Thermal conductivity                   | W/m·K               | W/m·K                 |
| Viscosity              | `MU`, `viscosity`        | Dynamic viscosity                      | Pa·s                | Pa·s                  |
| Prandtl Number         | `Prandtl`                | Prandtl number                         | (unitless)          | (unitless)            |
| Compressibility Factor | `Z`                      | Compressibility factor                 | (unitless)          | (unitless)            |
| Surface Tension        | `surface_tension`        | Surface tension                        | N/m                 | N/m                   |
| Molar Mass             | `MM`, `molar_mass`       | Molar mass                             | kg/kmol             | kg/kmol               |
| Critical Pressure      | `Pcrit`, `p_critical`    | Critical pressure                      | bar                 | Pa                    |
| Critical Temperature   | `Tcrit`                  | Critical temperature                   | °C                  | K                     |
| Triple Point Temp      | `Ttriple`                | Triple point temperature               | °C                  | K                     |
| Dew Point Temp         | `Tdp`, `dewpoint`        | Dew point temperature (humid air)      | °C                  | K                     |
| Wet Bulb Temp          | `Twb`, `wetbulb`         | Wet bulb temperature (humid air)       | °C                  | K                     |
| Enthalpy (dry air)     | `Hda`                    | Specific enthalpy (dry air base)       | kJ/kg               | J/kg                  |
| Enthalpy (humid air)   | `Hha`                    | Specific enthalpy (humid air)          | kJ/kg               | J/kg                  |
| Entropy (dry air)      | `Sda`                    | Specific entropy (dry air base)        | kJ/kg·K             | J/kg·K                |
| Entropy (humid air)    | `Sha`                    | Specific entropy (humid air)           | kJ/kg·K             | J/kg·K                |
| Relative Humidity      | `RH`, `relhum`, `R`      | Relative humidity (humid air)          | p.u.                | p.u.                  |
| Humidity Ratio         | `W`, `omega`, `humrat`   | Humidity ratio (humid air)             | kg_water/kg_dry_air | kg_water/kg_dry_air   |

For more details on available properties and fluids, refer to the [CoolProp documentation](http://www.coolprop.org/).

## If building from source:

Restore packages from the `.csproj`

```bash
dotnet restore
```

Build release

```bash
dotnet build -c Release
```

Move the output files:

```bash
.\bin\Release\net6.0-windows\publish\CoolPropWrapper64-packed.xll
.\bin\Release\net6.0-windows\CoolPropWrapper.dll
.\CoolProp.dll
```

Import in Excel the `.xll` file.
