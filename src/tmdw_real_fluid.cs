using System;
using System.Collections.Generic;
using ExcelDna.Integration;

public static partial class CoolPropWrapper
{
    // Function to calculate thermodynamic properties using SI units (no conversion)
    [ExcelFunction(Name = "PropsSI", Description = "Calculate thermodynamic properties of real fluids using CoolProp with SI units (K, Pa, J/kg, etc.). Supports array inputs for value1 and/or value2.")]
    public static object PropsSI(string output, string name1, object value1, string name2, object value2, string fluid)
    {
        if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
        if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
        if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
        if (string.IsNullOrWhiteSpace(fluid)) return "Error: Fluid name is missing.";
        if (value1 == null) return "Error: First property value is missing.";
        if (value2 == null) return "Error: Second property value is missing.";

        name1 = FormatName(name1);
        name2 = FormatName(name2);
        output = FormatName(output);
        fluid = FormatFluidName(fluid);

        bool isValue1Array = IsArrayInput(value1);
        bool isValue2Array = IsArrayInput(value2);

        if (!isValue1Array && !isValue2Array)
        {
            if (!(value1 is double)) return "Error: First property value is not a number.";
            if (!(value2 is double)) return "Error: Second property value is not a number.";

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

            if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
            {
                return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
            }

            return result;
        }

        // Handle array inputs
        try
        {
            double[] values1;
            double[] values2;

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

            if (isValue1Array && isValue2Array && values1.Length != values2.Length)
            {
                return $"Error: Array lengths must match. value1 has {values1.Length} elements, value2 has {values2.Length} elements.";
            }

            int resultLength = Math.Max(values1.Length, values2.Length);
            double[] results = new double[resultLength];

            for (int i = 0; i < resultLength; i++)
            {
                double val1 = values1[isValue1Array ? i : 0];
                double val2 = values2[isValue2Array ? i : 0];

                double result;
                try
                {
                    result = PropsSI_Internal(output, name1, val1, name2, val2, fluid);
                }
                catch (Exception ex)
                {
                    return $"Error at index {i}: {ex.Message}. CoolProp error: {GetCoolPropError()}";
                }

                if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
                {
                    return $"Error at index {i}: CoolProp failed to compute property. {GetCoolPropError()}";
                }

                results[i] = result;
            }

            bool outputAsRow = (isValue1Array && IsRowArray(value1)) || (isValue2Array && IsRowArray(value2));
            
            object[,] outputArray;
            if (outputAsRow)
            {
                outputArray = new object[1, resultLength];
                for (int i = 0; i < resultLength; i++)
                {
                    outputArray[0, i] = results[i];
                }
            }
            else
            {
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
        if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
        if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
        if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
        if (string.IsNullOrWhiteSpace(fluid)) return "Error: Fluid name is missing.";
        if (value1 == null) return "Error: First property value is missing.";
        if (value2 == null) return "Error: Second property value is missing.";

        name1 = FormatName(name1);
        name2 = FormatName(name2);
        output = FormatName(output);
        fluid = FormatFluidName(fluid);

        bool isValue1Array = IsArrayInput(value1);
        bool isValue2Array = IsArrayInput(value2);

        if (!isValue1Array && !isValue2Array)
        {
            if (!(value1 is double)) return "Error: First property value is not a number.";
            if (!(value2 is double)) return "Error: Second property value is not a number.";

            double val1SI = ConvertToSI(name1, (double)value1);
            double val2SI = ConvertToSI(name2, (double)value2);

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

            if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
            {
                return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
            }

            return ConvertFromSI(output, result);
        }

        // Handle array inputs
        try
        {
            double[] values1;
            double[] values2;

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

            if (isValue1Array && isValue2Array && values1.Length != values2.Length)
            {
                return $"Error: Array lengths must match. value1 has {values1.Length} elements, value2 has {values2.Length} elements.";
            }

            int resultLength = Math.Max(values1.Length, values2.Length);
            double[] results = new double[resultLength];

            for (int i = 0; i < resultLength; i++)
            {
                double val1 = values1[isValue1Array ? i : 0];
                double val2 = values2[isValue2Array ? i : 0];

                double val1SI = ConvertToSI(name1, val1);
                double val2SI = ConvertToSI(name2, val2);

                double result;
                try
                {
                    result = PropsSI_Internal(output, name1, val1SI, name2, val2SI, fluid);
                }
                catch (Exception ex)
                {
                    return $"Error at index {i}: {ex.Message}. CoolProp error: {GetCoolPropError()}";
                }

                if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
                {
                    return $"Error at index {i}: CoolProp failed to compute property. {GetCoolPropError()}";
                }

                results[i] = ConvertFromSI(output, result);
            }

            bool outputAsRow = (isValue1Array && IsRowArray(value1)) || (isValue2Array && IsRowArray(value2));
            
            object[,] outputArray;
            if (outputAsRow)
            {
                outputArray = new object[1, resultLength];
                for (int i = 0; i < resultLength; i++)
                {
                    outputArray[0, i] = results[i];
                }
            }
            else
            {
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

    // Function to get phase using SI units
    [ExcelFunction(Name = "PhaseSI", Description = "Get the phase of a fluid using CoolProp with SI units (K, Pa).")]
    public static object PhaseSI(string name1, object value1, string name2, object value2, string fluid)
    {
        if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
        if (value1 == null || !(value1 is double)) return "Error: First property value is missing or not a number.";
        if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
        if (value2 == null || !(value2 is double)) return "Error: Second property value is missing or not a number.";
        if (string.IsNullOrWhiteSpace(fluid)) return "Error: Fluid name is missing.";

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
        if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
        if (value1 == null || !(value1 is double)) return "Error: First property value is missing or not a number.";
        if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
        if (value2 == null || !(value2 is double)) return "Error: Second property value is missing or not a number.";
        if (string.IsNullOrWhiteSpace(fluid)) return "Error: Fluid name is missing.";

        name1 = FormatName(name1);
        name2 = FormatName(name2);
        fluid = FormatFluidName(fluid);

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
        if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
        if (string.IsNullOrWhiteSpace(fluid)) return "Error: Fluid name is missing.";

        output = FormatName(output);
        fluid = FormatFluidName(fluid);

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

        if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
        {
            return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
        }

        return result;
    }

    // Function to calculate single-input properties using engineering units
    [ExcelFunction(Name = "Props1", Description = "Calculate single-input properties (critical properties, molar mass, etc.) using CoolProp with engineering units (°C, bar, kg/mol, etc.).")]
    public static object Props1(string output, string fluid)
    {
        if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
        if (string.IsNullOrWhiteSpace(fluid)) return "Error: Fluid name is missing.";

        output = FormatName(output);
        fluid = FormatFluidName(fluid);

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

        if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
        {
            return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
        }

        return ConvertFromSI(output, result);
    }

    // Function to get CoolProp version and global parameters
    [ExcelFunction(Name = "GetGlobalParam", Description = "Get CoolProp version, revision, or other global parameter strings.")]
    public static object GetGlobalParam(string param)
    {
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
        if (string.IsNullOrWhiteSpace(fluid)) return "Error: Fluid name is missing.";
        if (string.IsNullOrWhiteSpace(param)) return "Error: Parameter name is missing.";

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
            if (elements == null || fractions == null)
                return "Error: Elements and fractions arrays cannot be null.";

            int elemCount = elements.GetLength(0);
            int fracCount = fractions.GetLength(0);

            if (elemCount != fracCount)
                return $"Error: Number of elements ({elemCount}) must match number of fractions ({fracCount}).";

            if (elemCount == 0)
                return "Error: At least one element is required.";

            var components = new List<string>();
            for (int i = 0; i < elemCount; i++)
            {
                object elemValue = elements.GetLength(1) == 1 ? elements[i, 0] : elements[0, i];
                if (elemValue == null || string.IsNullOrWhiteSpace(elemValue.ToString()))
                    continue;

                string element = elemValue.ToString().Trim();

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

            return "HEOS::" + string.Join("&", components);
        }
        catch (Exception ex)
        {
            return $"Error creating mixture string: {ex.Message}";
        }
    }
}
