using System;
using ExcelDna.Integration;

public static partial class CoolPropWrapper
{
    // Helper functions for array detection and extraction
    private static bool IsArrayInput(object input)
    {
        return input is object[,] || input is double[,];
    }

    /// <summary>
    /// Returns a descriptive error string when a numeric argument receives a non-numeric value.
    /// Distinguishes Excel error values (#N/A, #REF!, etc.) from other unexpected types.
    /// </summary>
    private static string DescribeNonNumericError(string paramName, object value)
    {
        if (value is ExcelError err)
            return $"Error: {paramName} contains an Excel error value ({err}). " +
                   "Ensure the referenced cell contains a valid number, not an error formula.";
        if (value is string s)
            return $"Error: {paramName} is a text value (\"{s}\"). A numeric value is required.";
        if (value is bool b)
            return $"Error: {paramName} is a boolean ({b}). A numeric value is required.";
        if (value == null || value is ExcelMissing || value is ExcelEmpty)
            return $"Error: {paramName} is empty or missing. A numeric value is required.";
        return $"Error: {paramName} is not a number (received {value.GetType().Name}: {value}). A numeric value is required.";
    }

    private static bool IsRowArray(object input)
    {
        if (input is object[,] arr2D)
        {
            return arr2D.GetLength(0) == 1 && arr2D.GetLength(1) > 1;
        }
        else if (input is double[,] dblArr2D)
        {
            return dblArr2D.GetLength(0) == 1 && dblArr2D.GetLength(1) > 1;
        }
        return false;
    }

    private static double[] ExtractDoubleArray(object input)
    {
        if (input is object[,] arr2D)
        {
            int rows = arr2D.GetLength(0);
            int cols = arr2D.GetLength(1);

            if (rows == 1) // Row array
            {
                var result = new double[cols];
                for (int i = 0; i < cols; i++)
                {
                    if (arr2D[0, i] is double d)
                        result[i] = d;
                    else
                        throw new InvalidCastException($"Element at position {i} is not a number.");
                }
                return result;
            }
            else if (cols == 1) // Column array
            {
                var result = new double[rows];
                for (int i = 0; i < rows; i++)
                {
                    if (arr2D[i, 0] is double d)
                        result[i] = d;
                    else
                        throw new InvalidCastException($"Element at position {i} is not a number.");
                }
                return result;
            }
            else
            {
                throw new ArgumentException("Array must be a single row or single column.");
            }
        }
        else if (input is double[,] dblArr2D)
        {
            int rows = dblArr2D.GetLength(0);
            int cols = dblArr2D.GetLength(1);

            if (rows == 1) // Row array
            {
                var result = new double[cols];
                for (int i = 0; i < cols; i++)
                    result[i] = dblArr2D[0, i];
                return result;
            }
            else if (cols == 1) // Column array
            {
                var result = new double[rows];
                for (int i = 0; i < rows; i++)
                    result[i] = dblArr2D[i, 0];
                return result;
            }
            else
            {
                throw new ArgumentException("Array must be a single row or single column.");
            }
        }

        throw new ArgumentException("Input is not a valid array type.");
    }

    // ---- Broadcasting helpers ----

    // Returns {rows, cols} shape of any parameter: scalar→{1,1}, object[,]→actual dims.
    internal static int[] GetParamShape(object param)
    {
        if (param is object[,] arr)  return new int[] { arr.GetLength(0),  arr.GetLength(1)  };
        if (param is double[,] darr) return new int[] { darr.GetLength(0), darr.GetLength(1) };
        return new int[] { 1, 1 };
    }

    // Try to extract a string at grid position (r, c) using broadcasting (scalars always return [0,0]).
    internal static bool TryGetStringAt(object param, int[] shape, int r, int c, out string value)
    {
        int ri = shape[0] == 1 ? 0 : r;
        int ci = shape[1] == 1 ? 0 : c;
        if (param is string s)       { value = s;    return true; }
        if (param is ExcelMissing || param is ExcelEmpty) { value = null; return false; }
        if (param is object[,] arr)
        {
            object cell = arr[ri, ci];
            if (cell is string cs)   { value = cs;   return true; }
            value = null; return false;
        }
        value = null; return false;
    }

    // Try to extract a double at grid position (r, c) using broadcasting.
    internal static bool TryGetDoubleAt(object param, int[] shape, int r, int c, out double value)
    {
        int ri = shape[0] == 1 ? 0 : r;
        int ci = shape[1] == 1 ? 0 : c;
        if (param is double d)       { value = d;    return true; }
        if (param is object[,] arr)
        {
            object cell = arr[ri, ci];
            if (cell is double dv)   { value = dv;   return true; }
            value = double.NaN; return false;
        }
        if (param is double[,] darr) { value = darr[ri, ci]; return true; }
        value = double.NaN; return false;
    }

    // Return the raw object at position (r, c) — used to pass to DescribeNonNumericError.
    internal static object GetRawAt(object param, int[] shape, int r, int c)
    {
        int ri = shape[0] == 1 ? 0 : r;
        int ci = shape[1] == 1 ? 0 : c;
        if (param is object[,] arr)  return arr[ri, ci];
        if (param is double[,] darr) return darr[ri, ci];
        return param;
    }

    // Resolve output grid size (M rows × N cols) from a set of parameter shapes using broadcasting rules:
    //   - Row arrays (rows=1) → contribute to column count N
    //   - Column arrays (cols=1) → contribute to row count M
    //   - 2-D arrays → contribute to both M and N
    // Returns false if two non-scalar extents conflict on the same axis.
    internal static bool ResolveOutputShape(out int M, out int N, params int[][] shapes)
    {
        M = 1; N = 1;
        foreach (var sh in shapes)
        {
            if (sh[0] > 1) { if (M > 1 && M != sh[0]) return false; M = sh[0]; }
            if (sh[1] > 1) { if (N > 1 && N != sh[1]) return false; N = sh[1]; }
        }
        return true;
    }
}
