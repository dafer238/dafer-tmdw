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
}
