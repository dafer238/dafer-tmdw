using System;
using System.Collections.Concurrent;
using System.Text;
using System.Threading;
using ExcelDna.Integration;

public static partial class CoolPropWrapper
{
    private static readonly ConcurrentDictionary<string, double> _propCache
        = new ConcurrentDictionary<string, double>(StringComparer.Ordinal);
    private static volatile bool _cacheEnabled = true;
    private static int _cacheHits = 0;
    private static int _cacheMisses = 0;

    // Thread-local error capture: captured immediately after a failed native call so that
    // a parallel CoolProp call on another thread cannot overwrite the global error string
    // before this thread reads it.
    [ThreadStatic]
    private static string _threadLastError;

    internal static string ConsumeLastError()
    {
        string e = _threadLastError;
        _threadLastError = null;
        return e ?? GetCoolPropError();
    }

    // Minimum array length that gets dispatched to the thread pool for parallel evaluation.
    internal const int ParallelThreshold = 4;

    // ---- Cached wrappers (SI inputs/outputs) ----

    private static double CachedPropsSI(string output, string n1, double v1, string n2, double v2, string fluid)
    {
        if (!_cacheEnabled)
            return PropsSI_Internal(output, n1, v1, n2, v2, fluid);

        string key = string.Concat("RF|", output, "|", n1, "|", v1.ToString("R"), "|", n2, "|", v2.ToString("R"), "|", fluid);
        if (_propCache.TryGetValue(key, out double hit))
        {
            Interlocked.Increment(ref _cacheHits);
            return hit;
        }
        double result = PropsSI_Internal(output, n1, v1, n2, v2, fluid);
        if (!double.IsNaN(result) && result < 1.0E+308 && result > -1.0E+308)
        {
            _propCache.TryAdd(key, result);
            Interlocked.Increment(ref _cacheMisses);
        }
        else
        {
            _threadLastError = GetCoolPropError();
        }
        return result;
    }

    private static double CachedHAPropsSI(string output, string n1, double v1, string n2, double v2, string n3, double v3)
    {
        if (!_cacheEnabled)
            return HAPropsSI_Internal(output, n1, v1, n2, v2, n3, v3);

        string key = string.Concat("HA|", output, "|", n1, "|", v1.ToString("R"), "|", n2, "|", v2.ToString("R"), "|", n3, "|", v3.ToString("R"));
        if (_propCache.TryGetValue(key, out double hit))
        {
            Interlocked.Increment(ref _cacheHits);
            return hit;
        }
        double result = HAPropsSI_Internal(output, n1, v1, n2, v2, n3, v3);
        if (!double.IsNaN(result) && result < 1.0E+308 && result > -1.0E+308)
        {
            _propCache.TryAdd(key, result);
            Interlocked.Increment(ref _cacheMisses);
        }
        else
        {
            _threadLastError = GetCoolPropError();
        }
        return result;
    }

    private static double CachedProps1SI(string output, string fluid)
    {
        if (!_cacheEnabled)
            return Props1SI_Internal(output, fluid);

        string key = string.Concat("P1|", output, "|", fluid);
        if (_propCache.TryGetValue(key, out double hit))
        {
            Interlocked.Increment(ref _cacheHits);
            return hit;
        }
        double result = Props1SI_Internal(output, fluid);
        if (!double.IsNaN(result) && result < 1.0E+308 && result > -1.0E+308)
        {
            _propCache.TryAdd(key, result);
            Interlocked.Increment(ref _cacheMisses);
        }
        return result;
    }

    private static double CachedGasMixSI(string output, string n1, double v1, string n2, double v2,
        int[] speciesIdx, double[] molarFracs)
    {
        if (!_cacheEnabled)
            return GasMixtureCalculator.CalcProperty(output, n1, v1, n2, v2, speciesIdx, molarFracs);

        var sb = new StringBuilder("GM|", 64);
        sb.Append(output).Append('|').Append(n1).Append('|').Append(v1.ToString("R"))
          .Append('|').Append(n2).Append('|').Append(v2.ToString("R")).Append('|');
        for (int k = 0; k < speciesIdx.Length; k++)
        {
            if (k > 0) sb.Append(',');
            sb.Append(speciesIdx[k]).Append(':').Append(molarFracs[k].ToString("R"));
        }
        string key = sb.ToString();

        if (_propCache.TryGetValue(key, out double hit))
        {
            Interlocked.Increment(ref _cacheHits);
            return hit;
        }
        double result = GasMixtureCalculator.CalcProperty(output, n1, v1, n2, v2, speciesIdx, molarFracs);
        if (!double.IsNaN(result))
        {
            _propCache.TryAdd(key, result);
            Interlocked.Increment(ref _cacheMisses);
        }
        return result;
    }

    // ---- Excel-exposed cache management functions ----

    [ExcelFunction(Name = "CP_ClearCache",
        Description = "Clears the CoolProp result cache and resets hit/miss counters. Returns the number of entries that were cleared.")]
    public static object CP_ClearCache()
    {
        int count = _propCache.Count;
        _propCache.Clear();
        Interlocked.Exchange(ref _cacheHits, 0);
        Interlocked.Exchange(ref _cacheMisses, 0);
        LogDebug($"Cache cleared: {count} entries removed.");
        return count;
    }

    [ExcelFunction(Name = "CP_CacheStats", IsVolatile = true,
        Description = "Returns cache statistics as a string: entry count, hits, misses, and enabled status. Recalculates automatically on every F9 / auto-calc cycle.")]
    public static object CP_CacheStats()
    {
        return $"Entries: {_propCache.Count} | Hits: {_cacheHits} | Misses: {_cacheMisses} | Enabled: {_cacheEnabled}";
    }

    [ExcelFunction(Name = "CP_EnableCache",
        Description = "Enable or disable the result cache. Cache is enabled by default. " +
                      "Pass TRUE to enable, FALSE to disable and clear all cached entries.")]
    public static object CP_EnableCache(bool enable)
    {
        _cacheEnabled = enable;
        if (!enable)
        {
            _propCache.Clear();
            Interlocked.Exchange(ref _cacheHits, 0);
            Interlocked.Exchange(ref _cacheMisses, 0);
        }
        LogDebug($"Cache {(enable ? "enabled" : "disabled and cleared")}.");
        return $"Cache {(enable ? "enabled" : "disabled and cleared")}.";
    }
}
