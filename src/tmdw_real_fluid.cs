using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelDna.Integration;

public static partial class CoolPropWrapper
{
    [ExcelFunction(Name = "PropsSI", Description = "Calculate thermodynamic properties of real fluids using CoolProp with SI units (K, Pa, J/kg). All six parameters (output, name1, name2, fluid, value1, value2) accept a scalar, a column array (M×1), a row array (1×N), or a 2-D range. Mixing a column and a row produces a Cartesian-product table; equal-size arrays iterate element-wise.")]
    public static object PropsSI(object output, object name1, object value1, object name2, object value2, object fluid)
    {
        int[] outSh = GetParamShape(output);
        int[] n1Sh  = GetParamShape(name1);
        int[] v1Sh  = GetParamShape(value1);
        int[] n2Sh  = GetParamShape(name2);
        int[] v2Sh  = GetParamShape(value2);
        int[] flSh  = GetParamShape(fluid);

        if (!ResolveOutputShape(out int M, out int N, outSh, n1Sh, v1Sh, n2Sh, v2Sh, flSh))
            return "Error: Incompatible array dimensions. Row arrays (1×N) must share the same N; column arrays (M×1) must share the same M.";

        if (M == 1 && N == 1)
        {
            if (!TryGetStringAt(output, outSh, 0, 0, out string outStr) || string.IsNullOrWhiteSpace(outStr))
                return "Error: Output parameter is missing or not a string.";
            if (!TryGetStringAt(name1, n1Sh, 0, 0, out string n1Str) || string.IsNullOrWhiteSpace(n1Str))
                return "Error: First property name is missing or not a string.";
            if (!TryGetStringAt(name2, n2Sh, 0, 0, out string n2Str) || string.IsNullOrWhiteSpace(n2Str))
                return "Error: Second property name is missing or not a string.";
            if (!TryGetStringAt(fluid, flSh, 0, 0, out string flStr) || string.IsNullOrWhiteSpace(flStr))
                return "Error: Fluid name is missing or not a string.";

            outStr = FormatName(outStr);
            n1Str  = FormatName(n1Str);
            n2Str  = FormatName(n2Str);
            flStr  = FormatFluidName(flStr);

            if (!TryGetDoubleAt(value1, v1Sh, 0, 0, out double v1)) return DescribeNonNumericError("value1", GetRawAt(value1, v1Sh, 0, 0));
            if (!TryGetDoubleAt(value2, v2Sh, 0, 0, out double v2)) return DescribeNonNumericError("value2", GetRawAt(value2, v2Sh, 0, 0));

            double result;
            try { result = CachedPropsSI(outStr, n1Str, v1, n2Str, v2, flStr); }
            catch (DllNotFoundException ex)
            {
                LogDebug($"PropsSI DllNotFoundException: {ex.Message}");
                return $"Error: CoolProp.dll not found. Use CPropDiag() to see paths checked. {ex.Message}";
            }
            catch (EntryPointNotFoundException ex)
            {
                LogDebug($"PropsSI EntryPointNotFoundException: {ex.Message}");
                return $"Error: Required CoolProp functions not found in DLL — check that CoolProp.dll is the 64-bit version. {ex.Message}";
            }
            catch (Exception ex)
            {
                LogDebug($"PropsSI exception ({outStr},{n1Str},{v1},{n2Str},{v2},{flStr}): {ex.Message}");
                return $"Error: Exception occurred while computing property. {ex.Message}. CoolProp error: {GetCoolPropError()}";
            }

            if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
            {
                string err = GetCoolPropError();
                LogDebug($"PropsSI invalid result ({outStr}): {err}");
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
            if (!TryGetStringAt(fluid, flSh, r, c, out string flStr) || string.IsNullOrWhiteSpace(flStr))
                { results[r, c] = "Error: Fluid name is empty at this position."; return; }

            outStr = FormatName(outStr);
            n1Str  = FormatName(n1Str);
            n2Str  = FormatName(n2Str);
            flStr  = FormatFluidName(flStr);

            if (!TryGetDoubleAt(value1, v1Sh, r, c, out double v1)) { results[r, c] = "Error: value1 is not a number at this position."; return; }
            if (!TryGetDoubleAt(value2, v2Sh, r, c, out double v2)) { results[r, c] = "Error: value2 is not a number at this position."; return; }

            try
            {
                double rv = CachedPropsSI(outStr, n1Str, v1, n2Str, v2, flStr);
                results[r, c] = (double.IsNaN(rv) || rv >= 1.0E+308 || rv <= -1.0E+308)
                    ? (object)$"Error: CoolProp failed: {GetCoolPropError()}"
                    : rv;
            }
            catch (Exception ex) { results[r, c] = $"Error: {ex.Message}"; }
        };

        if (total >= ParallelThreshold) Parallel.For(0, total, compute);
        else for (int i = 0; i < total; i++) compute(i);

        return results;
    }

    [ExcelFunction(Name = "Props", Description = "Calculate thermodynamic properties of real fluids using CoolProp with engineering units (°C, bar, kJ/kg). All six parameters (output, name1, name2, fluid, value1, value2) accept a scalar, a column array (M×1), a row array (1×N), or a 2-D range. Mixing a column and a row produces a Cartesian-product table; equal-size arrays iterate element-wise.")]
    public static object Props(object output, object name1, object value1, object name2, object value2, object fluid)
    {
        int[] outSh = GetParamShape(output);
        int[] n1Sh  = GetParamShape(name1);
        int[] v1Sh  = GetParamShape(value1);
        int[] n2Sh  = GetParamShape(name2);
        int[] v2Sh  = GetParamShape(value2);
        int[] flSh  = GetParamShape(fluid);

        if (!ResolveOutputShape(out int M, out int N, outSh, n1Sh, v1Sh, n2Sh, v2Sh, flSh))
            return "Error: Incompatible array dimensions. Row arrays (1×N) must share the same N; column arrays (M×1) must share the same M.";

        if (M == 1 && N == 1)
        {
            if (!TryGetStringAt(output, outSh, 0, 0, out string outStr) || string.IsNullOrWhiteSpace(outStr))
                return "Error: Output parameter is missing or not a string.";
            if (!TryGetStringAt(name1, n1Sh, 0, 0, out string n1Str) || string.IsNullOrWhiteSpace(n1Str))
                return "Error: First property name is missing or not a string.";
            if (!TryGetStringAt(name2, n2Sh, 0, 0, out string n2Str) || string.IsNullOrWhiteSpace(n2Str))
                return "Error: Second property name is missing or not a string.";
            if (!TryGetStringAt(fluid, flSh, 0, 0, out string flStr) || string.IsNullOrWhiteSpace(flStr))
                return "Error: Fluid name is missing or not a string.";

            outStr = FormatName(outStr);
            n1Str  = FormatName(n1Str);
            n2Str  = FormatName(n2Str);
            flStr  = FormatFluidName(flStr);

            if (!TryGetDoubleAt(value1, v1Sh, 0, 0, out double v1)) return DescribeNonNumericError("value1", GetRawAt(value1, v1Sh, 0, 0));
            if (!TryGetDoubleAt(value2, v2Sh, 0, 0, out double v2)) return DescribeNonNumericError("value2", GetRawAt(value2, v2Sh, 0, 0));

            double val1SI = ConvertToSI(n1Str, v1);
            double val2SI = ConvertToSI(n2Str, v2);

            double result;
            try { result = CachedPropsSI(outStr, n1Str, val1SI, n2Str, val2SI, flStr); }
            catch (DllNotFoundException ex)
            {
                LogDebug($"Props DllNotFoundException: {ex.Message}");
                return $"Error: CoolProp.dll not found in any search path. Use CPropDiag() function to see paths checked. {ex.Message}";
            }
            catch (EntryPointNotFoundException ex)
            {
                LogDebug($"Props EntryPointNotFoundException: {ex.Message}");
                return $"Error: Required functions not found in CoolProp.dll. Check DLL version and architecture match. {ex.Message}";
            }
            catch (Exception ex)
            {
                LogDebug($"Props exception ({outStr},{n1Str},{v1},{n2Str},{v2},{flStr}): {ex.Message}");
                return $"Error: Exception occurred while computing property. {ex.Message}. CoolProp error: {GetCoolPropError()}";
            }

            if (double.IsNaN(result) || result >= 1.0E+308 || result <= -1.0E+308)
            {
                string err = GetCoolPropError();
                LogDebug($"Props invalid result ({outStr}): {err}");
                return $"Error: CoolProp failed to compute property. {err}";
            }
            return ConvertFromSI(outStr, result);
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
            if (!TryGetStringAt(fluid, flSh, r, c, out string flStr) || string.IsNullOrWhiteSpace(flStr))
                { results[r, c] = "Error: Fluid name is empty at this position."; return; }

            outStr = FormatName(outStr);
            n1Str  = FormatName(n1Str);
            n2Str  = FormatName(n2Str);
            flStr  = FormatFluidName(flStr);

            if (!TryGetDoubleAt(value1, v1Sh, r, c, out double v1)) { results[r, c] = "Error: value1 is not a number at this position."; return; }
            if (!TryGetDoubleAt(value2, v2Sh, r, c, out double v2)) { results[r, c] = "Error: value2 is not a number at this position."; return; }

            double v1SI = ConvertToSI(n1Str, v1);
            double v2SI = ConvertToSI(n2Str, v2);

            try
            {
                double rv = CachedPropsSI(outStr, n1Str, v1SI, n2Str, v2SI, flStr);
                results[r, c] = (double.IsNaN(rv) || rv >= 1.0E+308 || rv <= -1.0E+308)
                    ? (object)$"Error: CoolProp failed: {GetCoolPropError()}"
                    : ConvertFromSI(outStr, rv);
            }
            catch (Exception ex) { results[r, c] = $"Error: {ex.Message}"; }
        };

        if (total >= ParallelThreshold) Parallel.For(0, total, compute);
        else for (int i = 0; i < total; i++) compute(i);

        return results;
    }

    [ExcelFunction(Name = "TMPr", Description = "Calculate thermodynamic properties of real fluids using CoolProp with engineering units (°C, bar, kJ/kg, etc.). Alias for Props.")]
    public static object TMPr(object output, object name1, object value1, object name2, object value2, object fluid)
    {
        return Props(output, name1, value1, name2, value2, fluid);
    }

    [ExcelFunction(Name = "PhaseSI", Description = "Get the phase of a fluid using CoolProp with SI units (K, Pa).")]
    public static object PhaseSI(string name1, object value1, string name2, object value2, string fluid)
    {
        if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
        if (value1 == null || !(value1 is double)) return DescribeNonNumericError("value1", value1);
        if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
        if (value2 == null || !(value2 is double)) return DescribeNonNumericError("value2", value2);
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

    [ExcelFunction(Name = "Phase", Description = "Get the phase of a fluid using CoolProp with engineering units (°C, bar).")]
    public static object Phase(string name1, object value1, string name2, object value2, string fluid)
    {
        if (string.IsNullOrWhiteSpace(name1)) return "Error: First property name is missing.";
        if (value1 == null || !(value1 is double)) return DescribeNonNumericError("value1", value1);
        if (string.IsNullOrWhiteSpace(name2)) return "Error: Second property name is missing.";
        if (value2 == null || !(value2 is double)) return DescribeNonNumericError("value2", value2);
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

    [ExcelFunction(Name = "Props1SI", Description = "Calculate single-input properties (critical properties, molar mass, etc.) using CoolProp with SI units (K, Pa, kg/mol, etc.).")]
    public static object Props1SI(string output, string fluid)
    {
        if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
        if (string.IsNullOrWhiteSpace(fluid)) return "Error: Fluid name is missing.";

        output = FormatName(output);
        fluid  = FormatFluidName(fluid);

        double result;
        try { result = CachedProps1SI(output, fluid); }
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
            return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";

        return result;
    }

    [ExcelFunction(Name = "Props1", Description = "Calculate single-input properties (critical properties, molar mass, etc.) using CoolProp with engineering units (°C, bar, kg/mol, etc.).")]
    public static object Props1(string output, string fluid)
    {
        if (string.IsNullOrWhiteSpace(output)) return "Error: Output parameter is missing.";
        if (string.IsNullOrWhiteSpace(fluid)) return "Error: Fluid name is missing.";

        output = FormatName(output);
        fluid  = FormatFluidName(fluid);

        double result;
        try { result = CachedProps1SI(output, fluid); }
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
            return $"Error: CoolProp failed to compute property. {GetCoolPropError()}";

        return ConvertFromSI(output, result);
    }

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

    // ---- Private helpers used by gas mixture functions ----

    private static double[] ResolveArrayOrScalar(object value, bool isArray, string paramName)
    {
        if (isArray) return ExtractDoubleArray(value);
        if (value is double d) return new double[] { d };
        return null;
    }

    private static object[,] BuildOutputArray(double[] results, int n,
        bool isValue1Array, object value1, bool isValue2Array, object value2)
    {
        bool asRow = (isValue1Array && IsRowArray(value1)) || (isValue2Array && IsRowArray(value2));
        var arr = asRow ? new object[1, n] : new object[n, 1];
        for (int i = 0; i < n; i++)
        {
            if (asRow) arr[0, i] = results[i];
            else       arr[i, 0] = results[i];
        }
        return arr;
    }
}
