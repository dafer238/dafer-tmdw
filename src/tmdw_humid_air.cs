using System;
using System.Threading.Tasks;
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
                result = CachedHAPropsSI(output, name1, (double)value1, name2, (double)value2, name3, (double)value3);
            }
            catch (DllNotFoundException ex)
            {
                LogDebug($"HAPropsSI DllNotFoundException: {ex.Message}");
                return $"Error: CoolProp.dll not found in any search path. Use CPropDiag() function to see paths checked. {ex.Message}";
            }
            catch (EntryPointNotFoundException ex)
            {
                LogDebug($"HAPropsSI EntryPointNotFoundException: {ex.Message}");
                return $"Error: Required functions not found in CoolProp.dll. Check DLL version and architecture match. {ex.Message}";
            }
            catch (Exception ex)
            {
                LogDebug($"HAPropsSI exception ({output}): {ex.Message}");
                return $"Error: Exception occurred while computing property. {ex.Message}. CoolProp error: {GetCoolPropError()}";
            }

            if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
            {
                string err = GetCoolPropError();
                LogDebug($"HAPropsSI invalid result ({output}): {err}");
                return $"Error: CoolProp failed to compute property. {err}";
            }

            return result;
        }

        try
        {
            double[] values1 = ResolveArrayOrScalar(value1, isValue1Array, "First");
            if (values1 == null) return "Error: First property value is not a number.";
            double[] values2 = ResolveArrayOrScalar(value2, isValue2Array, "Second");
            if (values2 == null) return "Error: Second property value is not a number.";
            double[] values3 = ResolveArrayOrScalar(value3, isValue3Array, "Third");
            if (values3 == null) return "Error: Third property value is not a number.";

            int n = Math.Max(Math.Max(values1.Length, values2.Length), values3.Length);
            if ((isValue1Array && values1.Length != n && values1.Length > 1) ||
                (isValue2Array && values2.Length != n && values2.Length > 1) ||
                (isValue3Array && values3.Length != n && values3.Length > 1))
            {
                return $"Error: All array inputs must have the same length. Lengths: value1={values1.Length}, value2={values2.Length}, value3={values3.Length}";
            }

            var results = new double[n];
            var errors = new string[n];

            Action<int> compute = i =>
            {
                double v1 = values1[isValue1Array ? Math.Min(i, values1.Length - 1) : 0];
                double v2 = values2[isValue2Array ? Math.Min(i, values2.Length - 1) : 0];
                double v3 = values3[isValue3Array ? Math.Min(i, values3.Length - 1) : 0];
                try
                {
                    double r = CachedHAPropsSI(output, name1, v1, name2, v2, name3, v3);
                    if (double.IsNaN(r) || r >= 1.0E+308 || r <= -1.0E+308)
                        errors[i] = $"Error at index {i}: CoolProp failed to compute property. {GetCoolPropError()}";
                    else
                        results[i] = r;
                }
                catch (Exception ex) { errors[i] = $"Error at index {i}: {ex.Message}"; }
            };

            if (n >= ParallelThreshold)
                Parallel.For(0, n, compute);
            else
                for (int i = 0; i < n; i++) compute(i);

            for (int i = 0; i < n; i++)
                if (errors[i] != null) { LogDebug($"HAPropsSI array: {errors[i]}"); return errors[i]; }

            return BuildOutputArray(results, n,
                isValue1Array, value1,
                isValue2Array || isValue3Array, isValue2Array ? value2 : value3);
        }
        catch (InvalidCastException ex) { return $"Error: {ex.Message}"; }
        catch (ArgumentException ex) { return $"Error: {ex.Message}"; }
        catch (Exception ex) { return $"Error processing array inputs: {ex.Message}"; }
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
                result = CachedHAPropsSI(output, name1, val1SI, name2, val2SI, name3, val3SI);
            }
            catch (DllNotFoundException ex)
            {
                LogDebug($"HAProps DllNotFoundException: {ex.Message}");
                return $"Error: CoolProp.dll not found in any search path. Use CPropDiag() function to see paths checked. {ex.Message}";
            }
            catch (EntryPointNotFoundException ex)
            {
                LogDebug($"HAProps EntryPointNotFoundException: {ex.Message}");
                return $"Error: Required functions not found in CoolProp.dll. Check DLL version and architecture match. {ex.Message}";
            }
            catch (Exception ex)
            {
                LogDebug($"HAProps exception ({output}): {ex.Message}");
                return $"Error: Exception occurred while computing property. {ex.Message}. CoolProp error: {GetCoolPropError()}";
            }

            if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
            {
                string err = GetCoolPropError();
                LogDebug($"HAProps invalid result ({output}): {err}");
                return $"Error: CoolProp failed to compute property. {err}";
            }

            return ConvertFromSI_HA(output, result);
        }

        try
        {
            double[] values1 = ResolveArrayOrScalar(value1, isValue1Array, "First");
            if (values1 == null) return "Error: First property value is not a number.";
            double[] values2 = ResolveArrayOrScalar(value2, isValue2Array, "Second");
            if (values2 == null) return "Error: Second property value is not a number.";
            double[] values3 = ResolveArrayOrScalar(value3, isValue3Array, "Third");
            if (values3 == null) return "Error: Third property value is not a number.";

            int n = Math.Max(Math.Max(values1.Length, values2.Length), values3.Length);
            if ((isValue1Array && values1.Length != n && values1.Length > 1) ||
                (isValue2Array && values2.Length != n && values2.Length > 1) ||
                (isValue3Array && values3.Length != n && values3.Length > 1))
            {
                return $"Error: All array inputs must have the same length. Lengths: value1={values1.Length}, value2={values2.Length}, value3={values3.Length}";
            }

            var results = new double[n];
            var errors = new string[n];

            Action<int> compute = i =>
            {
                double v1SI = ConvertToSI_HA(name1, values1[isValue1Array ? Math.Min(i, values1.Length - 1) : 0]);
                double v2SI = ConvertToSI_HA(name2, values2[isValue2Array ? Math.Min(i, values2.Length - 1) : 0]);
                double v3SI = ConvertToSI_HA(name3, values3[isValue3Array ? Math.Min(i, values3.Length - 1) : 0]);
                try
                {
                    double r = CachedHAPropsSI(output, name1, v1SI, name2, v2SI, name3, v3SI);
                    if (double.IsNaN(r) || r >= 1.0E+308 || r <= -1.0E+308)
                        errors[i] = $"Error at index {i}: CoolProp failed to compute property. {GetCoolPropError()}";
                    else
                        results[i] = ConvertFromSI_HA(output, r);
                }
                catch (Exception ex) { errors[i] = $"Error at index {i}: {ex.Message}"; }
            };

            if (n >= ParallelThreshold)
                Parallel.For(0, n, compute);
            else
                for (int i = 0; i < n; i++) compute(i);

            for (int i = 0; i < n; i++)
                if (errors[i] != null) { LogDebug($"HAProps array: {errors[i]}"); return errors[i]; }

            return BuildOutputArray(results, n,
                isValue1Array, value1,
                isValue2Array || isValue3Array, isValue2Array ? value2 : value3);
        }
        catch (InvalidCastException ex) { return $"Error: {ex.Message}"; }
        catch (ArgumentException ex) { return $"Error: {ex.Message}"; }
        catch (Exception ex) { return $"Error processing array inputs: {ex.Message}"; }
    }

    [ExcelFunction(Name = "TMPa", Description = "Calculate thermodynamic properties of humid air using CoolProp with engineering units (°C, bar, kJ/kg, etc.). Alias for HAProps.")]
    public static object TMPa(string output, string name1, object value1, string name2, object value2, string name3, object value3)
    {
        return HAProps(output, name1, value1, name2, value2, name3, value3);
    }
}
