using System;
using ExcelDna.Integration;

public static partial class CoolPropWrapper
{
    [ExcelFunction(Name = "HAPropsSI", Description = "Calculate thermodynamic properties of humid air using CoolProp with SI units (K, Pa, J/kg, etc.). Supports array inputs.")]
    public static object HAPropsSI(string output, string name1, object value1, string name2, object value2, string name3, object value3)
    {
        if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
        if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
        if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
        if (string.IsNullOrWhiteSpace(name3)) return "Error: Third property name is missing.";
        if (value1 == null) return "Error: First property value is missing.";
        if (value2 == null) return "Error: Second property value is missing.";
        if (value3 == null) return "Error: Third property value is missing.";

        name1 = FormatName_HA(name1);
        name2 = FormatName_HA(name2);
        name3 = FormatName_HA(name3);
        output = FormatName_HA(output);

        bool isValue1Array = IsArrayInput(value1);
        bool isValue2Array = IsArrayInput(value2);
        bool isValue3Array = IsArrayInput(value3);

        if (!isValue1Array && !isValue2Array && !isValue3Array)
        {
            if (!(value1 is double)) return "Error: First property value is not a number.";
            if (!(value2 is double)) return "Error: Second property value is not a number.";
            if (!(value3 is double)) return "Error: Third property value is not a number.";

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

            if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
            {
                return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
            }

            return result;
        }

        // Handle array inputs
        try
        {
            double[] values1, values2, values3;

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

            int resultLength = Math.Max(Math.Max(values1.Length, values2.Length), values3.Length);
            if ((isValue1Array && values1.Length != resultLength && values1.Length > 1) ||
                (isValue2Array && values2.Length != resultLength && values2.Length > 1) ||
                (isValue3Array && values3.Length != resultLength && values3.Length > 1))
            {
                return $"Error: All array inputs must have the same length. Lengths: value1={values1.Length}, value2={values2.Length}, value3={values3.Length}";
            }

            double[] results = new double[resultLength];

            for (int i = 0; i < resultLength; i++)
            {
                double val1 = values1[isValue1Array ? Math.Min(i, values1.Length - 1) : 0];
                double val2 = values2[isValue2Array ? Math.Min(i, values2.Length - 1) : 0];
                double val3 = values3[isValue3Array ? Math.Min(i, values3.Length - 1) : 0];

                double result;
                try
                {
                    result = HAPropsSI_Internal(output, name1, val1, name2, val2, name3, val3);
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

            bool outputAsRow = (isValue1Array && IsRowArray(value1)) || 
                              (isValue2Array && IsRowArray(value2)) || 
                              (isValue3Array && IsRowArray(value3));
            
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

    [ExcelFunction(Name = "HAProps", Description = "Calculate thermodynamic properties of humid air using CoolProp with engineering units (°C, bar, kJ/kg, etc.). Supports array inputs.")]
    public static object HAProps(string output, string name1, object value1, string name2, object value2, string name3, object value3)
    {
        if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
        if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
        if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
        if (string.IsNullOrWhiteSpace(name3)) return "Error: Third property name is missing.";
        if (value1 == null) return "Error: First property value is missing.";
        if (value2 == null) return "Error: Second property value is missing.";
        if (value3 == null) return "Error: Third property value is missing.";

        name1 = FormatName_HA(name1);
        name2 = FormatName_HA(name2);
        name3 = FormatName_HA(name3);
        output = FormatName_HA(output);

        bool isValue1Array = IsArrayInput(value1);
        bool isValue2Array = IsArrayInput(value2);
        bool isValue3Array = IsArrayInput(value3);

        if (!isValue1Array && !isValue2Array && !isValue3Array)
        {
            if (!(value1 is double)) return "Error: First property value is not a number.";
            if (!(value2 is double)) return "Error: Second property value is not a number.";
            if (!(value3 is double)) return "Error: Third property value is not a number.";

            double val1SI = ConvertToSI_HA(name1, (double)value1);
            double val2SI = ConvertToSI_HA(name2, (double)value2);
            double val3SI = ConvertToSI_HA(name3, (double)value3);

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

            if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
            {
                return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";
            }

            return ConvertFromSI_HA(output, result);
        }

        // Handle array inputs
        try
        {
            double[] values1, values2, values3;

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

            int resultLength = Math.Max(Math.Max(values1.Length, values2.Length), values3.Length);
            if ((isValue1Array && values1.Length != resultLength && values1.Length > 1) ||
                (isValue2Array && values2.Length != resultLength && values2.Length > 1) ||
                (isValue3Array && values3.Length != resultLength && values3.Length > 1))
            {
                return $"Error: All array inputs must have the same length. Lengths: value1={values1.Length}, value2={values2.Length}, value3={values3.Length}";
            }

            double[] results = new double[resultLength];

            for (int i = 0; i < resultLength; i++)
            {
                double val1 = values1[isValue1Array ? Math.Min(i, values1.Length - 1) : 0];
                double val2 = values2[isValue2Array ? Math.Min(i, values2.Length - 1) : 0];
                double val3 = values3[isValue3Array ? Math.Min(i, values3.Length - 1) : 0];

                double val1SI = ConvertToSI_HA(name1, val1);
                double val2SI = ConvertToSI_HA(name2, val2);
                double val3SI = ConvertToSI_HA(name3, val3);

                double result;
                try
                {
                    result = HAPropsSI_Internal(output, name1, val1SI, name2, val2SI, name3, val3SI);
                }
                catch (Exception ex)
                {
                    return $"Error at index {i}: {ex.Message}. CoolProp error: {GetCoolPropError()}";
                }

                if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
                {
                    return $"Error at index {i}: CoolProp failed to compute property. {GetCoolPropError()}";
                }

                results[i] = ConvertFromSI_HA(output, result);
            }

            bool outputAsRow = (isValue1Array && IsRowArray(value1)) || 
                              (isValue2Array && IsRowArray(value2)) || 
                              (isValue3Array && IsRowArray(value3));
            
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

    // TMPa alias for HAProps (uses engineering units)
    [ExcelFunction(Name = "TMPa", Description = "Calculate thermodynamic properties of humid air using CoolProp with engineering units (°C, bar, kJ/kg, etc.). Alias for HAProps.")]
    public static object TMPa(string output, string name1, object value1, string name2, object value2, string name3, object value3)
    {
        return HAProps(output, name1, value1, name2, value2, name3, value3);
    }
}
