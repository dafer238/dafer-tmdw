using System;

public static partial class CoolPropWrapper
{
    // Helper functions for array detection and extraction
    private static bool IsArrayInput(object input)
    {
        return input is object[,] || input is double[,];
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
