using System;
using System.Collections.Generic;
using ExcelDna.Integration;

public static partial class CoolPropWrapper
{
    /// <summary>
    /// Gas mixture property name mapping. Reuses the existing PropertyNameMap where applicable,
    /// plus gas-mixture-specific aliases.
    /// </summary>
    private static readonly Dictionary<string, string> GasMixPropertyMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        // Temperature
        ["T"] = "T", ["TEMP"] = "T", ["TEMPERATURE"] = "T",
        // Pressure
        ["P"] = "P", ["PRES"] = "P", ["PRESSURE"] = "P",
        // Heat capacity
        ["CP"] = "Cpmass", ["CPMASS"] = "Cpmass", ["C"] = "Cpmass",
        ["CV"] = "Cvmass", ["CVMASS"] = "Cvmass",
        ["CPMOLAR"] = "Cpmolar", ["CPMOL"] = "Cpmolar",
        ["CVMOLAR"] = "Cvmolar", ["CVMOL"] = "Cvmolar",
        // Enthalpy
        ["H"] = "H", ["ENTH"] = "H", ["ENTHALPY"] = "H", ["HMASS"] = "H", ["H_MASS"] = "H",
        ["HMOLAR"] = "Hmolar", ["H_MOLAR"] = "Hmolar",
        // Entropy
        ["S"] = "S", ["ENTR"] = "S", ["ENTROPY"] = "S", ["SMASS"] = "S", ["S_MASS"] = "S",
        ["SMOLAR"] = "Smolar", ["S_MOLAR"] = "Smolar",
        // Internal energy
        ["U"] = "U", ["INTERNALENERGY"] = "U", ["UMASS"] = "U",
        // Density
        ["D"] = "D", ["RHO"] = "D", ["DENS"] = "D", ["DENSITY"] = "D", ["DMASS"] = "D", ["RHOMASS"] = "D",
        // Gamma
        ["GAMMA"] = "isentropic_expansion_coefficient", ["ISENTROPIC_EXP"] = "isentropic_expansion_coefficient",
        ["ISENTROPICEXPANSIONCOEFFICIENT"] = "isentropic_expansion_coefficient",
        // Speed of sound
        ["A"] = "A", ["W"] = "A", ["SPEED_OF_SOUND"] = "A", ["SPEEDOFSOUND"] = "A",
        // Molar mass
        ["MM"] = "M", ["MOLAR_MASS"] = "M", ["MOLARMASS"] = "M", ["MOLEMASS"] = "M",
        // Transport
        ["V"] = "viscosity", ["MU"] = "viscosity", ["VISCOSITY"] = "viscosity",
        ["K"] = "conductivity", ["CONDUCTIVITY"] = "conductivity", ["L"] = "conductivity",
        ["PRANDTL"] = "Prandtl", ["PR"] = "Prandtl",
        // Compressibility
        ["Z"] = "Z", ["COMPRESSIBILITY"] = "Z", ["COMPRESSIBILITYFACTOR"] = "Z"
    };

    /// <summary>Resolve gas mixture property name.</summary>
    private static string FormatGasMixName(string name)
    {
        if (GasMixPropertyMap.TryGetValue(name, out string mapped))
            return mapped;
        return name;
    }

    // Gas mixture energy properties (for unit conversion kJ <-> J)
    private static readonly HashSet<string> GasMixEnergyProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "H", "Hmolar", "U", "S", "Smolar", "Cpmass", "Cvmass", "Cpmolar", "Cvmolar"
    };

    /// <summary>Convert gas mixture input from engineering to SI.</summary>
    private static double ConvertToSI_GasMix(string name, double value)
    {
        if (name == "T") return value + 273.15;
        if (name == "P") return value * 1e5;
        if (GasMixEnergyProperties.Contains(name)) return value * 1000;
        return value;
    }

    /// <summary>Convert gas mixture output from SI to engineering.</summary>
    private static double ConvertFromSI_GasMix(string name, double value)
    {
        if (name == "T") return value - 273.15;
        if (name == "P") return value / 1e5;
        if (GasMixEnergyProperties.Contains(name)) return value / 1000;
        return value;
    }

    /// <summary>
    /// Parse gas names and fractions from Excel ranges, resolve species, normalize fractions.
    /// Returns null on error, sets errorMsg.
    /// </summary>
    private static bool ParseGasMixInputs(
        object gasNames, object fractions, object fractionBasis,
        out int[] speciesIndices, out double[] molarFractions, out string errorMsg)
    {
        speciesIndices = null;
        molarFractions = null;
        errorMsg = null;

        // Extract gas names
        string[] names;
        if (gasNames is object[,] nameArr)
        {
            int rows = nameArr.GetLength(0);
            int cols = nameArr.GetLength(1);
            int count = Math.Max(rows, cols);
            bool isRow = rows == 1;
            names = new string[count];
            for (int i = 0; i < count; i++)
            {
                object val = isRow ? nameArr[0, i] : nameArr[i, 0];
                if (val == null || string.IsNullOrWhiteSpace(val.ToString()))
                {
                    errorMsg = $"Error: Gas name at position {i + 1} is empty.";
                    return false;
                }
                names[i] = val.ToString().Trim();
            }
        }
        else if (gasNames is string singleName)
        {
            names = new string[] { singleName };
        }
        else
        {
            errorMsg = "Error: gasNames must be a string or array of strings.";
            return false;
        }

        // Extract fractions
        double[] fracs;
        if (fractions is object[,] fracArr)
        {
            int rows = fracArr.GetLength(0);
            int cols = fracArr.GetLength(1);
            int count = Math.Max(rows, cols);
            bool isRow = rows == 1;
            fracs = new double[count];
            for (int i = 0; i < count; i++)
            {
                object val = isRow ? fracArr[0, i] : fracArr[i, 0];
                if (val is double d)
                    fracs[i] = d;
                else
                {
                    errorMsg = $"Error: Fraction at position {i + 1} is not a number.";
                    return false;
                }
            }
        }
        else if (fractions is double singleFrac)
        {
            fracs = new double[] { singleFrac };
        }
        else
        {
            errorMsg = "Error: fractions must be a number or array of numbers.";
            return false;
        }

        if (names.Length != fracs.Length)
        {
            errorMsg = $"Error: Number of gas names ({names.Length}) must match number of fractions ({fracs.Length}).";
            return false;
        }

        // Resolve species
        speciesIndices = new int[names.Length];
        for (int i = 0; i < names.Length; i++)
        {
            int idx = GasSpeciesData.ResolveSpecies(names[i]);
            if (idx < 0)
            {
                errorMsg = $"Error: Unknown gas species '{names[i]}'. Check spelling or use formula (e.g., N2, CO2, H2O).";
                return false;
            }
            speciesIndices[i] = idx;
        }

        // Normalize fractions
        if (!GasMixtureCalculator.NormalizeFractions(fracs))
        {
            errorMsg = "Error: All fractions must be non-negative and at least one must be positive.";
            return false;
        }

        // Convert to molar fractions if needed
        string basis = "molar"; // default
        if (fractionBasis is string basisStr && !string.IsNullOrWhiteSpace(basisStr))
            basis = basisStr.Trim();

        if (string.Equals(basis, "mass", StringComparison.OrdinalIgnoreCase))
        {
            molarFractions = GasMixtureCalculator.MassToMolarFractions(fracs, speciesIndices);
        }
        else
        {
            molarFractions = fracs;
        }

        return true;
    }

    /// <summary>
    /// Core gas mixture calculation (SI units). Routes through cache; used by both SI and engineering-unit functions.
    /// </summary>
    private static double CalcGasMixSI(string output, string name1, double value1, string name2, double value2,
        int[] speciesIndices, double[] molarFractions)
    {
        return CachedGasMixSI(output, name1, value1, name2, value2, speciesIndices, molarFractions);
    }

    // ========== Excel-exposed functions ==========

    [ExcelFunction(Name = "PropsSIGasMix",
        Description = "Calculate thermodynamic/transport properties of ideal gas mixtures using SI units (K, Pa, J/kg). Supports NASA polynomial data for 21 species.")]
    public static object PropsSIGasMix(
        string output, string name1, object value1, string name2, object value2,
        object gasNames, object fractions,
        [ExcelArgument(Description = "Fraction basis: \"molar\" (default) or \"mass\"")] object fractionBasis)
    {
        if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
        if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
        if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
        if (value1 == null) return "Error: First property value is missing.";
        if (value2 == null) return "Error: Second property value is missing.";
        if (gasNames == null) return "Error: Gas names are missing.";
        if (fractions == null) return "Error: Fractions are missing.";

        // Parse gas mixture inputs
        if (!ParseGasMixInputs(gasNames, fractions, fractionBasis,
            out int[] speciesIndices, out double[] molarFractions, out string parseError))
        {
            return parseError;
        }

        // Format property names
        output = FormatGasMixName(output);
        name1 = FormatGasMixName(name1);
        name2 = FormatGasMixName(name2);

        bool isValue1Array = IsArrayInput(value1);
        bool isValue2Array = IsArrayInput(value2);

        // Single value calculation
        if (!isValue1Array && !isValue2Array)
        {
            if (!(value1 is double)) return "Error: First property value is not a number.";
            if (!(value2 is double)) return "Error: Second property value is not a number.";

            double result = CalcGasMixSI(output, name1, (double)value1, name2, (double)value2, speciesIndices, molarFractions);
            if (double.IsNaN(result))
                return "Error: Failed to compute gas mixture property. Check input pair and temperature range (200-6000 K).";
            return result;
        }

        // Array calculation
        try
        {
            double[] values1, values2;

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

            if (isValue1Array && isValue2Array && values1.Length != values2.Length)
                return $"Error: Array lengths must match. value1 has {values1.Length} elements, value2 has {values2.Length} elements.";

            int resultLength = Math.Max(values1.Length, values2.Length);
            double[] results = new double[resultLength];

            for (int i = 0; i < resultLength; i++)
            {
                double v1 = values1[isValue1Array ? i : 0];
                double v2 = values2[isValue2Array ? i : 0];
                double result = CalcGasMixSI(output, name1, v1, name2, v2, speciesIndices, molarFractions);
                if (double.IsNaN(result))
                    return $"Error at index {i}: Failed to compute gas mixture property.";
                results[i] = result;
            }

            bool outputAsRow = (isValue1Array && IsRowArray(value1)) || (isValue2Array && IsRowArray(value2));
            object[,] outputArray;
            if (outputAsRow)
            {
                outputArray = new object[1, resultLength];
                for (int i = 0; i < resultLength; i++)
                    outputArray[0, i] = results[i];
            }
            else
            {
                outputArray = new object[resultLength, 1];
                for (int i = 0; i < resultLength; i++)
                    outputArray[i, 0] = results[i];
            }
            return outputArray;
        }
        catch (Exception ex)
        {
            return $"Error processing array inputs: {ex.Message}";
        }
    }

    [ExcelFunction(Name = "PropsGasMix",
        Description = "Calculate thermodynamic/transport properties of ideal gas mixtures using engineering units (°C, bar, kJ/kg). Supports NASA polynomial data for 21 species.")]
    public static object PropsGasMix(
        string output, string name1, object value1, string name2, object value2,
        object gasNames, object fractions,
        [ExcelArgument(Description = "Fraction basis: \"molar\" (default) or \"mass\"")] object fractionBasis)
    {
        if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
        if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
        if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
        if (value1 == null) return "Error: First property value is missing.";
        if (value2 == null) return "Error: Second property value is missing.";
        if (gasNames == null) return "Error: Gas names are missing.";
        if (fractions == null) return "Error: Fractions are missing.";

        if (!ParseGasMixInputs(gasNames, fractions, fractionBasis,
            out int[] speciesIndices, out double[] molarFractions, out string parseError))
        {
            return parseError;
        }

        output = FormatGasMixName(output);
        name1 = FormatGasMixName(name1);
        name2 = FormatGasMixName(name2);

        bool isValue1Array = IsArrayInput(value1);
        bool isValue2Array = IsArrayInput(value2);

        if (!isValue1Array && !isValue2Array)
        {
            if (!(value1 is double)) return "Error: First property value is not a number.";
            if (!(value2 is double)) return "Error: Second property value is not a number.";

            double v1SI = ConvertToSI_GasMix(name1, (double)value1);
            double v2SI = ConvertToSI_GasMix(name2, (double)value2);

            double result = CalcGasMixSI(output, name1, v1SI, name2, v2SI, speciesIndices, molarFractions);
            if (double.IsNaN(result))
                return "Error: Failed to compute gas mixture property. Check input pair and temperature range (-73 to 5727 °C).";
            return ConvertFromSI_GasMix(output, result);
        }

        try
        {
            double[] values1, values2;

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

            if (isValue1Array && isValue2Array && values1.Length != values2.Length)
                return $"Error: Array lengths must match. value1 has {values1.Length} elements, value2 has {values2.Length} elements.";

            int resultLength = Math.Max(values1.Length, values2.Length);
            double[] results = new double[resultLength];

            for (int i = 0; i < resultLength; i++)
            {
                double v1 = values1[isValue1Array ? i : 0];
                double v2 = values2[isValue2Array ? i : 0];

                double v1SI = ConvertToSI_GasMix(name1, v1);
                double v2SI = ConvertToSI_GasMix(name2, v2);

                double result = CalcGasMixSI(output, name1, v1SI, name2, v2SI, speciesIndices, molarFractions);
                if (double.IsNaN(result))
                    return $"Error at index {i}: Failed to compute gas mixture property.";
                results[i] = ConvertFromSI_GasMix(output, result);
            }

            bool outputAsRow = (isValue1Array && IsRowArray(value1)) || (isValue2Array && IsRowArray(value2));
            object[,] outputArray;
            if (outputAsRow)
            {
                outputArray = new object[1, resultLength];
                for (int i = 0; i < resultLength; i++)
                    outputArray[0, i] = results[i];
            }
            else
            {
                outputArray = new object[resultLength, 1];
                for (int i = 0; i < resultLength; i++)
                    outputArray[i, 0] = results[i];
            }
            return outputArray;
        }
        catch (Exception ex)
        {
            return $"Error processing array inputs: {ex.Message}";
        }
    }

    // TMPg alias for PropsGasMix (uses engineering units)
    [ExcelFunction(Name = "TMPg",
        Description = "Calculate thermodynamic/transport properties of ideal gas mixtures using engineering units (°C, bar, kJ/kg). Alias for PropsGasMix.")]
    public static object TMPg(
        string output, string name1, object value1, string name2, object value2,
        object gasNames, object fractions,
        [ExcelArgument(Description = "Fraction basis: \"molar\" (default) or \"mass\"")] object fractionBasis)
    {
        return PropsGasMix(output, name1, value1, name2, value2, gasNames, fractions, fractionBasis);
    }
}
