# Plan: Gas Mixture Properties + Source Code Refactor

## Problem Statement

1. CoolProp's real-gas mixture calculations diverge at high temperatures (gas turbine exhaust). Need a custom function for multi-component ideal gas mixtures.
2. The entire codebase is in a single 1700-line `CoolPropWrapper.cs` file. Need to split into logical files for maintainability.

## Proposed Approach

### New Feature: Gas Mixture Properties
Use **NASA 7-coefficient polynomials** + **ideal gas mixing rules** — industry standard for combustion gases, valid 200-6000K. Standalone — no CoolProp dependency.

### Refactor: Split Source Files
Split the monolithic `CoolPropWrapper.cs` into partial classes and separate static classes. The .csproj SDK-style project auto-includes all .cs files, so no build config changes needed.

#### New file structure:
```
src/
├── tmdw.cs                     (partial) DLL loading, diagnostics, static constructor
├── tmdw_mappings.cs            (partial) All dictionary/hashset static field definitions
├── tmdw_conversions.cs         (partial) ConvertToSI, ConvertFromSI, FormatName, FormatFluidName
├── tmdw_helpers.cs             (partial) IsArrayInput, ExtractDoubleArray, IsRowArray
├── tmdw_real_fluid.cs          (partial) Props, PropsSI, Props1, Props1SI, Phase, PhaseSI, TMPr, MixtureString
├── tmdw_humid_air.cs           (partial) HAProps, HAPropsSI, TMPa
├── gas_mixture/
│   ├── gas_species_data.cs     Static data: NASA coefficients, molar masses, transport params, name aliases
│   ├── gas_mixture_calc.cs     Core logic: thermo, transport, mixing, state solver, fraction handling
│   └── tmdw_gas_mixture.cs     (partial CoolPropWrapper) Excel-exposed functions: PropsGasMix, PropsSIGasMix, TMPg
├── CoolPropWrapper.csproj
└── CoolPropWrapper.dna
```

## Technical Details

### Thermodynamic Properties (NASA Polynomials)

For each species, using two temperature ranges (200-1000K and 1000-6000K):

```
Cp/R = a1 + a2*T + a3*T² + a4*T³ + a5*T⁴
H/(RT) = a1 + a2*T/2 + a3*T²/3 + a4*T³/4 + a5*T⁴/5 + a6/T
S/R = a1*ln(T) + a2*T + a3*T²/2 + a4*T³/3 + a5*T⁴/4 + a7
```

### Ideal Gas Mixing Rules

- `Cp_mix = Σ(yi * Cp_i)` (molar basis, then convert to mass basis)
- `H_mix = Σ(yi * H_i)` (molar basis)
- `S_mix = Σ(yi * S_i) - R * Σ(yi * ln(yi))` (includes entropy of mixing)
- `Cv = Cp - R` (per mole), `γ = Cp/Cv`
- `ρ = P * M_mix / (R * T)` (ideal gas law)
- `a = √(γ * R_specific * T)` (speed of sound)

### Transport Properties

- **Viscosity**: Sutherland's law for individual species + Wilke's mixing rule
- **Thermal conductivity**: Sutherland's law + Wassiljewa-Mason-Saxena mixing rule
- **Prandtl number**: Pr = μ * Cp / k

### Supported Species (with NASA polynomial data from GRI-Mech 3.0)

| Gas | Formula | Molar Mass (g/mol) | T Range |
|-----|---------|-------------------|---------|
| Nitrogen | N2 | 28.014 | 300-5000K |
| Oxygen | O2 | 31.998 | 200-3500K |
| Carbon Dioxide | CO2 | 44.009 | 200-3500K |
| Water Vapor | H2O | 18.015 | 200-3500K |
| Argon | Ar | 39.948 | 300-5000K |
| Carbon Monoxide | CO | 28.010 | 200-3500K |
| Hydrogen | H2 | 2.016 | 200-3500K |
| Nitric Oxide | NO | 30.006 | 200-6000K |
| Nitrogen Dioxide | NO2 | 46.006 | 200-6000K |
| Methane | CH4 | 16.043 | 200-3500K |
| Nitrous Oxide | N2O | 44.013 | 200-6000K |
| Sulfur Dioxide | SO2 | 64.066 | 200-6000K |
| Helium | He | 4.003 | 200-6000K |
| Neon | Ne | 20.180 | 200-6000K |
| Ammonia | NH3 | 17.031 | 200-6000K |

but also add common components present in natural gas like ethane, propane, butane, etc.

### Supported Output Properties

| Property | Name(s) | Description | Units (Eng) | Units (SI) |
|----------|---------|-------------|-------------|------------|
| Cp | CP, CPMASS | Mass-specific heat at constant P | kJ/(kg·K) | J/(kg·K) |
| Cv | CV, CVMASS | Mass-specific heat at constant V | kJ/(kg·K) | J/(kg·K) |
| H | ENTHALPY, HMASS | Mass-specific enthalpy | kJ/kg | J/kg |
| S | ENTROPY, SMASS | Mass-specific entropy | kJ/(kg·K) | J/(kg·K) |
| U | INTERNALENERGY | Mass-specific internal energy | kJ/kg | J/kg |
| D | DENSITY, RHO | Density (ideal gas law) | kg/m³ | kg/m³ |
| T | TEMPERATURE | Temperature | °C | K |
| P | PRESSURE | Pressure | bar | Pa |
| Gamma | ISENTROPIC_EXP | Cp/Cv ratio | - | - |
| A | SPEED_OF_SOUND | Speed of sound | m/s | m/s |
| MM | MOLAR_MASS | Mixture molar mass | kg/mol | kg/mol |
| V | VISCOSITY, MU | Dynamic viscosity | Pa·s | Pa·s |
| K | CONDUCTIVITY | Thermal conductivity | W/(m·K) | W/(m·K) |
| Prandtl | PRANDTL | Prandtl number | - | - |
| Z | COMPRESSIBILITY | Compressibility (always ~1) | - | - |
| Cpmolar | CPMOLAR | Molar heat at constant P | kJ/(kmol·K) | J/(mol·K) |
| Cvmolar | CVMOLAR | Molar heat at constant V | kJ/(kmol·K) | J/(mol·K) |
| Hmolar | HMOLAR | Molar enthalpy | kJ/kmol | J/mol |
| Smolar | SMOLAR | Molar entropy | kJ/(kmol·K) | J/(mol·K) |

### Supported State Input Pairs

| Input 1 | Input 2 | Method |
|---------|---------|--------|
| T | P | Direct calculation |
| P | T | Direct calculation |
| H | P | Newton-Raphson iteration to find T |
| P | H | Newton-Raphson iteration to find T |
| S | P | Newton-Raphson iteration to find T |
| P | S | Newton-Raphson iteration to find T |

### Function Signatures

#### Engineering Units (°C, bar, kJ/kg)
```
=PropsGasMix(output, name1, value1, name2, value2, gasNames, fractions, fractionBasis)
=TMPg(...)  // Alias
```

#### SI Units (K, Pa, J/kg)
```
=PropsSIGasMix(output, name1, value1, name2, value2, gasNames, fractions, fractionBasis)
```

**Parameters:**
- `output`: Property to calculate (e.g., "H", "Cp", "S", "D")
- `name1`: First state property name (e.g., "T", "P", "H")
- `value1`: First state property value (scalar or array)
- `name2`: Second state property name
- `value2`: Second state property value (scalar or array)
- `gasNames`: Array of gas names (e.g., {"N2", "O2", "CO2", "H2O"})
- `fractions`: Array of fractions (e.g., {0.75, 0.12, 0.06, 0.07})
- `fractionBasis`: "molar" (default) or "mass"

**Fraction normalization:** If fractions don't sum to 1, each fraction is proportionally scaled: `fi_normalized = fi / Σfj`

**Array support:** value1 and/or value2 can be Excel arrays. If arrays, output is an array of the same shape (row/column preserved).

## Implementation Todos

### 0. `refactor-split-files` — Refactor: Split CoolPropWrapper.cs into multiple files
Split the monolithic CoolPropWrapper.cs (1697 lines) into partial class files (all lowercase, underscore-separated):
- `tmdw.cs` → DLL loading, diagnostics, static constructor, internal wrappers
- `tmdw_mappings.cs` → PropertyNameMap, FluidNameMap, HumidAirPropertyMap, TemperatureProperties, PressureProperties, EnergyProperties
- `tmdw_conversions.cs` → ConvertToSI, ConvertFromSI, ConvertToSI_HA, ConvertFromSI_HA, FormatName, FormatFluidName, FormatName_HA
- `tmdw_helpers.cs` → IsArrayInput, IsRowArray, ExtractDoubleArray
- `tmdw_real_fluid.cs` → Props, PropsSI, Props1, Props1SI, Phase, PhaseSI, TMPr, MixtureString, GetGlobalParam, GetFluidParam
- `tmdw_humid_air.cs` → HAProps, HAPropsSI, TMPa
Delete old CoolPropWrapper.cs after split. Verify build succeeds.

### 1. `gas-species-data` — Define gas species data structures
Create static data structures containing:
- NASA polynomial coefficients (high-T and low-T) for all 15 species
- Molar masses
- Sutherland's law parameters (μ_ref, T_ref, S_μ) for viscosity
- Sutherland's law parameters (k_ref, T_ref, S_k) for conductivity
- Gas name alias mappings (e.g., "N2" → "Nitrogen", "OXYGEN" → "O2")

### 2. `species-thermo-calc` — Implement individual species property calculation
Functions to compute Cp, H, S for a single species at a given temperature T (in K) using NASA polynomials. Select correct coefficient set based on T range.

### 3. `mixture-thermo-calc` — Implement mixture property calculations
Combine individual species properties using ideal gas mixing rules:
- Weighted sums for Cp, H, S (molar then convert to mass-specific)
- Ideal gas law for density
- Derived properties: Cv, U, gamma, speed of sound, Z

### 4. `transport-properties` — Implement transport property calculations
- Sutherland's law for individual species viscosity and conductivity
- Wilke's mixing rule for mixture viscosity
- Wassiljewa-Mason-Saxena for mixture thermal conductivity
- Prandtl number from Pr = μ*Cp/k

### 5. `fraction-handling` — Implement fraction normalization and conversion
- Normalize fractions to sum to 1 (proportional scaling)
- Convert mass fractions ↔ molar fractions
- Validate inputs (no negatives, at least one component)

### 6. `state-solver` — Implement iterative state solver
- Direct calculation when T and P are both given
- Newton-Raphson iteration for H+P → T and S+P → T pairs
- Convergence criteria and max iterations with error handling

### 7. `excel-functions` — Implement Excel-exposed functions
Three Excel functions following existing patterns:
- `PropsSIGasMix` — SI units (K, Pa, J/kg)
- `PropsGasMix` — Engineering units (°C, bar, kJ/kg)
- `TMPg` — Alias for GasMixProps
With full array support (row/column detection, matching existing patterns).

### 8. `build-and-verify` — Build and verify compilation
Run `build.bat` to ensure everything compiles without errors.

## Dependencies
- Todo 0 has no dependencies (refactor first)
- Todo 1 depends on 0
- Todo 2 depends on 1
- Todo 3 depends on 2
- Todo 4 depends on 1
- Todo 5 depends on 1
- Todo 6 depends on 3
- Todo 7 depends on 3, 4, 5, 6
- Todo 8 depends on 7
- Todo 9 (commit + release) depends on 8

## Notes
- Use `partial class CoolPropWrapper` for all files that contain CoolPropWrapper members
- GasSpeciesData and GasMixtureCalculator are separate static classes (not partial CoolPropWrapper)
- GasMixtureFunctions.cs uses partial CoolPropWrapper to expose Excel functions
- SDK-style .csproj auto-includes all .cs files — no project file changes needed
- No new NuGet packages needed — pure C# math calculations
- NASA polynomial data sourced from GRI-Mech 3.0 thermodynamic database
- Commit all changes and create release tag v0.3.0
