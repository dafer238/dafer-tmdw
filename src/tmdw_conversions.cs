using System;

public static partial class CoolPropWrapper
{
    // Optimized property name mapping using dictionary lookups
    private static string FormatName(string name)
    {
        if (PropertyNameMap.TryGetValue(name, out string mapped))
            return mapped;
        return name;
    }

    // Optimize fluid name mapping using dictionary lookups
    private static string FormatFluidName(string fluid)
    {
        // Handle mixture strings (contain "&" or "HEOS::")
        if (fluid.Contains("&") || fluid.StartsWith("HEOS::", StringComparison.OrdinalIgnoreCase))
            return fluid;

        // Normalize by removing dashes, underscores, and spaces
        string normalized = fluid.Replace("-", "").Replace("_", "").Replace(" ", "");
        
        if (FluidNameMap.TryGetValue(normalized, out string mapped))
            return mapped;
        
        return fluid;
    }

    // Optimized humid air property name mapping
    private static string FormatName_HA(string name)
    {
        if (HumidAirPropertyMap.TryGetValue(name, out string mapped))
            return mapped;
        return name;
    }

    // Optimized unit conversion to SI using HashSet lookups
    private static double ConvertToSI(string name, double value)
    {
        if (TemperatureProperties.Contains(name))
            return value + 273.15;
        if (PressureProperties.Contains(name))
            return value * 1e5;
        if (EnergyProperties.Contains(name))
            return value * 1000;
        return value;
    }

    // Optimized unit conversion from SI using HashSet lookups
    private static double ConvertFromSI(string name, double value)
    {
        if (TemperatureProperties.Contains(name))
            return value - 273.15;
        if (PressureProperties.Contains(name))
            return value / 1e5;
        if (EnergyProperties.Contains(name))
            return value / 1000;
        return value;
    }

    // Convert from custom units to SI units for HA properties
    private static double ConvertToSI_HA(string name, double value)
    {
        if (TemperatureProperties.Contains(name))
            return value + 273.15;
        if (PressureProperties.Contains(name))
            return value * 1e5;
        if (EnergyProperties.Contains(name))
            return value * 1000;
        return value;
    }

    // Convert from SI units to custom units for HA properties
    private static double ConvertFromSI_HA(string name, double value)
    {
        if (TemperatureProperties.Contains(name))
            return value - 273.15;
        if (PressureProperties.Contains(name))
            return value / 1e5;
        if (EnergyProperties.Contains(name))
            return value / 1000;
        return value;
    }
}
