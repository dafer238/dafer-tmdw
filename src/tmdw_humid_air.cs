using System;
using System.Threading.Tasks;
using ExcelDna.Integration;

public static partial class CoolPropWrapper
{
    [ExcelFunction(Name = "HAPropsSI", Description = "Calculate thermodynamic properties of humid air using CoolProp with SI units (K, Pa, J/kg). All seven parameters (output, name1, name2, name3, value1, value2, value3) accept a scalar, a column array (M×1), a row array (1×N), or a 2-D range. Mixing a column and a row produces a Cartesian-product table; equal-size arrays iterate element-wise.")]
    public static object HAPropsSI(object output, object name1, object value1, object name2, object value2, object name3, object value3)
    {
        int[] outSh = GetParamShape(output);
        int[] n1Sh  = GetParamShape(name1);
        int[] v1Sh  = GetParamShape(value1);
        int[] n2Sh  = GetParamShape(name2);
        int[] v2Sh  = GetParamShape(value2);
        int[] n3Sh  = GetParamShape(name3);
        int[] v3Sh  = GetParamShape(value3);

        if (!ResolveOutputShape(out int M, out int N, outSh, n1Sh, v1Sh, n2Sh, v2Sh, n3Sh, v3Sh))
            return "Error: Incompatible array dimensions. Row arrays (1×N) must share the same N; column arrays (M×1) must share the same M.";

        if (M == 1 && N == 1)
        {
            if (!TryGetStringAt(output, outSh, 0, 0, out string outStr) || string.IsNullOrWhiteSpace(outStr))
                return "Error: Output parameter is missing or not a string.";
            if (!TryGetStringAt(name1, n1Sh, 0, 0, out string n1Str) || string.IsNullOrWhiteSpace(n1Str))
                return "Error: First property name is missing or not a string.";
            if (!TryGetStringAt(name2, n2Sh, 0, 0, out string n2Str) || string.IsNullOrWhiteSpace(n2Str))
                return "Error: Second property name is missing or not a string.";
            if (!TryGetStringAt(name3, n3Sh, 0, 0, out string n3Str) || string.IsNullOrWhiteSpace(n3Str))
                return "Error: Third property name is missing or not a string.";

            outStr = FormatName_HA(outStr);
            n1Str  = FormatName_HA(n1Str);
            n2Str  = FormatName_HA(n2Str);
            n3Str  = FormatName_HA(n3Str);

            if (!TryGetDoubleAt(value1, v1Sh, 0, 0, out double v1)) return DescribeNonNumericError("value1", GetRawAt(value1, v1Sh, 0, 0));
            if (!TryGetDoubleAt(value2, v2Sh, 0, 0, out double v2)) return DescribeNonNumericError("value2", GetRawAt(value2, v2Sh, 0, 0));
            if (!TryGetDoubleAt(value3, v3Sh, 0, 0, out double v3)) return DescribeNonNumericError("value3", GetRawAt(value3, v3Sh, 0, 0));

            double result;
            try { result = CachedHAPropsSI(outStr, n1Str, v1, n2Str, v2, n3Str, v3); }
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
                LogDebug($"HAPropsSI exception ({outStr}): {ex.Message}");
                return $"Error: Exception occurred while computing property. {ex.Message}. CoolProp error: {GetCoolPropError()}";
            }

            if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
            {
                string err = ConsumeLastError();
                LogDebug($"HAPropsSI invalid result ({outStr}): {err}");
                return $"Error: CoolProp failed to compute property. {err}";
            }
            return result;
        }

        var results = new object[M, N];
        int total   = M * N;

        Action<int> compute = idx =>
        {
            int r = idx / N, c = idx % N;
            if (!TryGetStringAt(output, outSh, r, c, out string outStr) || string.IsNullOrWhiteSpace(outStr))
                { results[r, c] = "Error: Output property name is empty at this position."; return; }
            if (!TryGetStringAt(name1, n1Sh, r, c, out string n1Str) || string.IsNullOrWhiteSpace(n1Str))
                { results[r, c] = "Error: First property name is empty at this position."; return; }
            if (!TryGetStringAt(name2, n2Sh, r, c, out string n2Str) || string.IsNullOrWhiteSpace(n2Str))
                { results[r, c] = "Error: Second property name is empty at this position."; return; }
            if (!TryGetStringAt(name3, n3Sh, r, c, out string n3Str) || string.IsNullOrWhiteSpace(n3Str))
                { results[r, c] = "Error: Third property name is empty at this position."; return; }

            outStr = FormatName_HA(outStr);
            n1Str  = FormatName_HA(n1Str);
            n2Str  = FormatName_HA(n2Str);
            n3Str  = FormatName_HA(n3Str);

            if (!TryGetDoubleAt(value1, v1Sh, r, c, out double v1)) { results[r, c] = "Error: value1 is not a number at this position."; return; }
            if (!TryGetDoubleAt(value2, v2Sh, r, c, out double v2)) { results[r, c] = "Error: value2 is not a number at this position."; return; }
            if (!TryGetDoubleAt(value3, v3Sh, r, c, out double v3)) { results[r, c] = "Error: value3 is not a number at this position."; return; }

            try
            {
                double rv = CachedHAPropsSI(outStr, n1Str, v1, n2Str, v2, n3Str, v3);
                results[r, c] = (double.IsNaN(rv) || rv >= 1.0E+308 || rv <= -1.0E+308)
                    ? (object)$"Error: CoolProp failed: {ConsumeLastError()}"
                    : rv;
            }
            catch (Exception ex) { results[r, c] = $"Error: {ex.Message}"; }
        };

        if (total >= ParallelThreshold) Parallel.For(0, total, compute);
        else for (int i = 0; i < total; i++) compute(i);

        return results;
    }

    [ExcelFunction(Name = "HAProps", Description = "Calculate thermodynamic properties of humid air using CoolProp with engineering units (°C, bar, kJ/kg). All seven parameters (output, name1, name2, name3, value1, value2, value3) accept a scalar, a column array (M×1), a row array (1×N), or a 2-D range. Mixing a column and a row produces a Cartesian-product table; equal-size arrays iterate element-wise.")]
    public static object HAProps(object output, object name1, object value1, object name2, object value2, object name3, object value3)
    {
        int[] outSh = GetParamShape(output);
        int[] n1Sh  = GetParamShape(name1);
        int[] v1Sh  = GetParamShape(value1);
        int[] n2Sh  = GetParamShape(name2);
        int[] v2Sh  = GetParamShape(value2);
        int[] n3Sh  = GetParamShape(name3);
        int[] v3Sh  = GetParamShape(value3);

        if (!ResolveOutputShape(out int M, out int N, outSh, n1Sh, v1Sh, n2Sh, v2Sh, n3Sh, v3Sh))
            return "Error: Incompatible array dimensions. Row arrays (1×N) must share the same N; column arrays (M×1) must share the same M.";

        if (M == 1 && N == 1)
        {
            if (!TryGetStringAt(output, outSh, 0, 0, out string outStr) || string.IsNullOrWhiteSpace(outStr))
                return "Error: Output parameter is missing or not a string.";
            if (!TryGetStringAt(name1, n1Sh, 0, 0, out string n1Str) || string.IsNullOrWhiteSpace(n1Str))
                return "Error: First property name is missing or not a string.";
            if (!TryGetStringAt(name2, n2Sh, 0, 0, out string n2Str) || string.IsNullOrWhiteSpace(n2Str))
                return "Error: Second property name is missing or not a string.";
            if (!TryGetStringAt(name3, n3Sh, 0, 0, out string n3Str) || string.IsNullOrWhiteSpace(n3Str))
                return "Error: Third property name is missing or not a string.";

            outStr = FormatName_HA(outStr);
            n1Str  = FormatName_HA(n1Str);
            n2Str  = FormatName_HA(n2Str);
            n3Str  = FormatName_HA(n3Str);

            if (!TryGetDoubleAt(value1, v1Sh, 0, 0, out double v1)) return DescribeNonNumericError("value1", GetRawAt(value1, v1Sh, 0, 0));
            if (!TryGetDoubleAt(value2, v2Sh, 0, 0, out double v2)) return DescribeNonNumericError("value2", GetRawAt(value2, v2Sh, 0, 0));
            if (!TryGetDoubleAt(value3, v3Sh, 0, 0, out double v3)) return DescribeNonNumericError("value3", GetRawAt(value3, v3Sh, 0, 0));

            double val1SI = ConvertToSI_HA(n1Str, v1);
            double val2SI = ConvertToSI_HA(n2Str, v2);
            double val3SI = ConvertToSI_HA(n3Str, v3);

            double result;
            try { result = CachedHAPropsSI(outStr, n1Str, val1SI, n2Str, val2SI, n3Str, val3SI); }
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
                LogDebug($"HAProps exception ({outStr}): {ex.Message}");
                return $"Error: Exception occurred while computing property. {ex.Message}. CoolProp error: {GetCoolPropError()}";
            }

            if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
            {
                string err = ConsumeLastError();
                LogDebug($"HAProps invalid result ({outStr}): {err}");
                return $"Error: CoolProp failed to compute property. {err}";
            }
            return ConvertFromSI_HA(outStr, result);
        }

        var results = new object[M, N];
        int total   = M * N;

        Action<int> compute = idx =>
        {
            int r = idx / N, c = idx % N;
            if (!TryGetStringAt(output, outSh, r, c, out string outStr) || string.IsNullOrWhiteSpace(outStr))
                { results[r, c] = "Error: Output property name is empty at this position."; return; }
            if (!TryGetStringAt(name1, n1Sh, r, c, out string n1Str) || string.IsNullOrWhiteSpace(n1Str))
                { results[r, c] = "Error: First property name is empty at this position."; return; }
            if (!TryGetStringAt(name2, n2Sh, r, c, out string n2Str) || string.IsNullOrWhiteSpace(n2Str))
                { results[r, c] = "Error: Second property name is empty at this position."; return; }
            if (!TryGetStringAt(name3, n3Sh, r, c, out string n3Str) || string.IsNullOrWhiteSpace(n3Str))
                { results[r, c] = "Error: Third property name is empty at this position."; return; }

            outStr = FormatName_HA(outStr);
            n1Str  = FormatName_HA(n1Str);
            n2Str  = FormatName_HA(n2Str);
            n3Str  = FormatName_HA(n3Str);

            if (!TryGetDoubleAt(value1, v1Sh, r, c, out double v1)) { results[r, c] = "Error: value1 is not a number at this position."; return; }
            if (!TryGetDoubleAt(value2, v2Sh, r, c, out double v2)) { results[r, c] = "Error: value2 is not a number at this position."; return; }
            if (!TryGetDoubleAt(value3, v3Sh, r, c, out double v3)) { results[r, c] = "Error: value3 is not a number at this position."; return; }

            double v1SI = ConvertToSI_HA(n1Str, v1);
            double v2SI = ConvertToSI_HA(n2Str, v2);
            double v3SI = ConvertToSI_HA(n3Str, v3);

            try
            {
                double rv = CachedHAPropsSI(outStr, n1Str, v1SI, n2Str, v2SI, n3Str, v3SI);
                results[r, c] = (double.IsNaN(rv) || rv >= 1.0E+308 || rv <= -1.0E+308)
                    ? (object)$"Error: CoolProp failed: {ConsumeLastError()}"
                    : ConvertFromSI_HA(outStr, rv);
            }
            catch (Exception ex) { results[r, c] = $"Error: {ex.Message}"; }
        };

        if (total >= ParallelThreshold) Parallel.For(0, total, compute);
        else for (int i = 0; i < total; i++) compute(i);

        return results;
    }

    [ExcelFunction(Name = "TMPa", Description = "Calculate thermodynamic properties of humid air using CoolProp with engineering units (°C, bar, kJ/kg, etc.). Alias for HAProps.")]
    public static object TMPa(object output, object name1, object value1, object name2, object value2, object name3, object value3)
    {
        return HAProps(output, name1, value1, name2, value2, name3, value3);
    }
}
