using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;

public static partial class CoolPropWrapper
{
    [ExcelFunction(Name = "CP_ListFluids",
        Description = "Returns a spilled column array of all pure fluid names available in the loaded CoolProp library.")]
    public static object CP_ListFluids()
    {
        try
        {
            string raw = get_global_param_string_buffer("fluids_list");
            if (string.IsNullOrEmpty(raw))
                return "Error: Could not retrieve fluid list from CoolProp. Ensure CoolProp.dll is loaded (use CPropDiag()).";

            string[] fluids = raw.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            object[,] result = new object[fluids.Length, 1];
            for (int i = 0; i < fluids.Length; i++)
                result[i, 0] = fluids[i].Trim();
            return result;
        }
        catch (Exception ex)
        {
            return $"Error listing fluids: {ex.Message}";
        }
    }

    [ExcelFunction(Name = "CP_ListProperties",
        Description = "Returns a spilled column array of all canonical CoolProp property names supported by Props/PropsSI.")]
    public static object CP_ListProperties()
    {
        try
        {
            var unique = new SortedSet<string>(PropertyNameMap.Values, StringComparer.OrdinalIgnoreCase);
            object[,] result = new object[unique.Count, 1];
            int i = 0;
            foreach (string s in unique)
                result[i++, 0] = s;
            return result;
        }
        catch (Exception ex)
        {
            return $"Error listing properties: {ex.Message}";
        }
    }

    [ExcelFunction(Name = "CP_ListPropertyAliases",
        Description = "Returns a spilled two-column array of all accepted property name aliases and their canonical CoolProp names, for Props/PropsSI.")]
    public static object CP_ListPropertyAliases()
    {
        try
        {
            var pairs = PropertyNameMap
                .OrderBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            object[,] result = new object[pairs.Length, 2];
            for (int i = 0; i < pairs.Length; i++)
            {
                result[i, 0] = pairs[i].Key;
                result[i, 1] = pairs[i].Value;
            }
            return result;
        }
        catch (Exception ex)
        {
            return $"Error listing property aliases: {ex.Message}";
        }
    }

    [ExcelFunction(Name = "CP_ListHAProperties",
        Description = "Returns a spilled column array of all canonical property names supported by HAProps/HAPropsSI.")]
    public static object CP_ListHAProperties()
    {
        try
        {
            var unique = new SortedSet<string>(HumidAirPropertyMap.Values, StringComparer.OrdinalIgnoreCase);
            object[,] result = new object[unique.Count, 1];
            int i = 0;
            foreach (string s in unique)
                result[i++, 0] = s;
            return result;
        }
        catch (Exception ex)
        {
            return $"Error listing HA properties: {ex.Message}";
        }
    }

    [ExcelFunction(Name = "CP_ListFluidAliases",
        Description = "Returns a spilled two-column array of all accepted fluid name aliases and their canonical CoolProp names.")]
    public static object CP_ListFluidAliases()
    {
        try
        {
            var pairs = FluidNameMap
                .OrderBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            object[,] result = new object[pairs.Length, 2];
            for (int i = 0; i < pairs.Length; i++)
            {
                result[i, 0] = pairs[i].Key;
                result[i, 1] = pairs[i].Value;
            }
            return result;
        }
        catch (Exception ex)
        {
            return $"Error listing fluid aliases: {ex.Message}";
        }
    }
}
