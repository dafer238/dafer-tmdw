using System;

/// <summary>
/// Core calculation engine for ideal gas mixtures.
/// All thermodynamic properties from NASA 7-coefficient polynomials.
/// Transport properties via Sutherland's law with Wilke/Wassiljewa mixing.
/// </summary>
public static class GasMixtureCalculator
{
    private const double R = GasSpeciesData.R; // 8.314462618 J/(mol·K)

    /// <summary>Result struct containing all computed mixture properties.</summary>
    public struct MixtureState
    {
        // Thermodynamic (mass-specific)
        public double T;          // Temperature (K)
        public double P;          // Pressure (Pa)
        public double CpMass;     // J/(kg·K)
        public double CvMass;     // J/(kg·K)
        public double HMass;      // J/kg
        public double SMass;      // J/(kg·K)
        public double UMass;      // J/kg
        public double Density;    // kg/m³
        public double Gamma;      // Cp/Cv
        public double SpeedOfSound; // m/s
        public double MolarMass;  // kg/mol
        public double Z;          // Compressibility factor (≈1 for ideal gas)
        // Thermodynamic (molar)
        public double CpMolar;    // J/(mol·K)
        public double CvMolar;    // J/(mol·K)
        public double HMolar;     // J/mol
        public double SMolar;     // J/(mol·K)
        // Transport
        public double Viscosity;      // Pa·s
        public double Conductivity;   // W/(m·K)
        public double Prandtl;        // dimensionless
    }

    #region Species-level thermodynamic calculations

    /// <summary>Cp/R for a single species at temperature T (K).</summary>
    private static double CpOverR(ref GasSpeciesData.GasSpecies sp, double T)
    {
        ref GasSpeciesData.NasaCoeffs c = ref (T <= sp.TMid ? ref sp.LowCoeffs : ref sp.HighCoeffs);
        return c.A1 + T * (c.A2 + T * (c.A3 + T * (c.A4 + T * c.A5)));
    }

    /// <summary>H/(RT) for a single species at temperature T (K).</summary>
    private static double HOverRT(ref GasSpeciesData.GasSpecies sp, double T)
    {
        ref GasSpeciesData.NasaCoeffs c = ref (T <= sp.TMid ? ref sp.LowCoeffs : ref sp.HighCoeffs);
        return c.A1 + T * (c.A2 * 0.5 + T * (c.A3 / 3.0 + T * (c.A4 * 0.25 + T * c.A5 * 0.2))) + c.A6 / T;
    }

    /// <summary>S/R for a single species at temperature T (K).</summary>
    private static double SOverR(ref GasSpeciesData.GasSpecies sp, double T)
    {
        ref GasSpeciesData.NasaCoeffs c = ref (T <= sp.TMid ? ref sp.LowCoeffs : ref sp.HighCoeffs);
        return c.A1 * Math.Log(T) + T * (c.A2 + T * (c.A3 * 0.5 + T * (c.A4 / 3.0 + T * c.A5 * 0.25))) + c.A7;
    }

    /// <summary>Cp in J/(mol·K) for a single species.</summary>
    public static double SpeciesCpMolar(int speciesIndex, double T)
    {
        return CpOverR(ref GasSpeciesData.Species[speciesIndex], T) * R;
    }

    /// <summary>H in J/mol for a single species.</summary>
    public static double SpeciesHMolar(int speciesIndex, double T)
    {
        return HOverRT(ref GasSpeciesData.Species[speciesIndex], T) * R * T;
    }

    /// <summary>S in J/(mol·K) for a single species at 1 atm reference.</summary>
    public static double SpeciesSMolar(int speciesIndex, double T)
    {
        return SOverR(ref GasSpeciesData.Species[speciesIndex], T) * R;
    }

    #endregion

    #region Transport property calculations

    /// <summary>Sutherland's law: property(T) = ref * (T/T_ref)^1.5 * (T_ref + S) / (T + S)</summary>
    private static double SutherlandValue(ref GasSpeciesData.SutherlandParams p, double T)
    {
        double ratio = T / p.TRef;
        return p.RefValue * ratio * Math.Sqrt(ratio) * (p.TRef + p.S) / (T + p.S);
    }

    /// <summary>
    /// Compute Wilke mixing rule phi_ij factor.
    /// phi_ij = (1/√8) * (1 + Mi/Mj)^(-0.5) * [1 + (μi/μj)^0.5 * (Mj/Mi)^0.25]^2
    /// </summary>
    private static double WilkePhi(double mu_i, double mu_j, double M_i, double M_j)
    {
        double muRatio = Math.Sqrt(mu_i / mu_j);
        double mRatio = Math.Pow(M_j / M_i, 0.25);
        double bracket = 1.0 + muRatio * mRatio;
        return bracket * bracket / Math.Sqrt(8.0 * (1.0 + M_i / M_j));
    }

    #endregion

    #region Fraction handling

    /// <summary>
    /// Normalize fractions so they sum to 1. Modifies array in place.
    /// Returns false if all fractions are zero or negative.
    /// </summary>
    public static bool NormalizeFractions(double[] fractions)
    {
        double sum = 0.0;
        for (int i = 0; i < fractions.Length; i++)
        {
            if (fractions[i] < 0.0) return false;
            sum += fractions[i];
        }
        if (sum <= 0.0) return false;
        if (Math.Abs(sum - 1.0) > 1e-12)
        {
            double invSum = 1.0 / sum;
            for (int i = 0; i < fractions.Length; i++)
                fractions[i] *= invSum;
        }
        return true;
    }

    /// <summary>
    /// Convert mass fractions to molar fractions.
    /// wi/Mi proportional to yi → yi = (wi/Mi) / Σ(wj/Mj)
    /// </summary>
    public static double[] MassToMolarFractions(double[] massFractions, int[] speciesIndices)
    {
        int n = massFractions.Length;
        double[] molar = new double[n];
        double sum = 0.0;
        for (int i = 0; i < n; i++)
        {
            molar[i] = massFractions[i] / GasSpeciesData.Species[speciesIndices[i]].MolarMass;
            sum += molar[i];
        }
        double invSum = 1.0 / sum;
        for (int i = 0; i < n; i++)
            molar[i] *= invSum;
        return molar;
    }

    /// <summary>
    /// Convert molar fractions to mass fractions.
    /// wi = yi*Mi / Σ(yj*Mj)
    /// </summary>
    public static double[] MolarToMassFractions(double[] molarFractions, int[] speciesIndices)
    {
        int n = molarFractions.Length;
        double[] mass = new double[n];
        double sum = 0.0;
        for (int i = 0; i < n; i++)
        {
            mass[i] = molarFractions[i] * GasSpeciesData.Species[speciesIndices[i]].MolarMass;
            sum += mass[i];
        }
        double invSum = 1.0 / sum;
        for (int i = 0; i < n; i++)
            mass[i] *= invSum;
        return mass;
    }

    #endregion

    #region Mixture property calculation

    /// <summary>
    /// Compute all mixture properties at given T (K) and P (Pa).
    /// speciesIndices: indices into GasSpeciesData.Species array.
    /// molarFractions: must be normalized and same length as speciesIndices.
    /// </summary>
    public static MixtureState CalcProperties(double T, double P, int[] speciesIndices, double[] molarFractions)
    {
        int n = speciesIndices.Length;
        var state = new MixtureState { T = T, P = P };

        // --- Mixture molar mass ---
        double Mmix = 0.0;
        for (int i = 0; i < n; i++)
            Mmix += molarFractions[i] * GasSpeciesData.Species[speciesIndices[i]].MolarMass;
        state.MolarMass = Mmix;

        // --- Thermodynamic properties (molar basis) ---
        double cpMolar = 0.0, hMolar = 0.0, sMolar = 0.0;
        for (int i = 0; i < n; i++)
        {
            double yi = molarFractions[i];
            if (yi <= 0.0) continue;
            int idx = speciesIndices[i];
            cpMolar += yi * CpOverR(ref GasSpeciesData.Species[idx], T);
            hMolar  += yi * HOverRT(ref GasSpeciesData.Species[idx], T);
            // Entropy of mixing: S_mix includes -R * Σ(yi * ln(yi))
            sMolar  += yi * SOverR(ref GasSpeciesData.Species[idx], T) - yi * Math.Log(yi);
        }
        // Convert from dimensionless to J/(mol·K) or J/mol
        state.CpMolar = cpMolar * R;
        state.HMolar = hMolar * R * T;
        state.CvMolar = state.CpMolar - R;  // Ideal gas: Cv = Cp - R (molar)
        state.SMolar = sMolar * R;

        // --- Convert to mass-specific ---
        double invM = 1.0 / Mmix;
        state.CpMass = state.CpMolar * invM;
        state.CvMass = state.CvMolar * invM;
        state.HMass = state.HMolar * invM;
        state.SMass = state.SMolar * invM;

        // --- Derived properties ---
        state.Gamma = state.CpMolar / state.CvMolar;
        double Rspecific = R * invM;  // J/(kg·K)
        state.Density = P * Mmix / (R * T);  // ρ = PM/(RT)
        state.UMass = state.HMass - P / state.Density;  // U = H - Pv = H - P/ρ
        state.SpeedOfSound = Math.Sqrt(state.Gamma * Rspecific * T);
        state.Z = 1.0;  // Ideal gas

        // --- Transport properties ---
        CalcTransportProperties(T, n, speciesIndices, molarFractions, ref state);

        return state;
    }

    /// <summary>Compute transport properties and add to state.</summary>
    private static void CalcTransportProperties(double T, int n, int[] speciesIndices, double[] molarFractions, ref MixtureState state)
    {
        // Compute individual species viscosity and conductivity
        double[] mu = new double[n];
        double[] k = new double[n];
        double[] M = new double[n];

        for (int i = 0; i < n; i++)
        {
            ref var sp = ref GasSpeciesData.Species[speciesIndices[i]];
            mu[i] = SutherlandValue(ref sp.ViscParams, T);
            k[i] = SutherlandValue(ref sp.CondParams, T);
            M[i] = sp.MolarMass;
        }

        // Wilke's mixing rule for viscosity
        // μ_mix = Σ(yi * μi / Σ(yj * φij))
        double muMix = 0.0;
        double kMix = 0.0;

        for (int i = 0; i < n; i++)
        {
            double yi = molarFractions[i];
            if (yi <= 0.0) continue;

            double phiSum = 0.0;
            for (int j = 0; j < n; j++)
            {
                double yj = molarFractions[j];
                if (yj <= 0.0) continue;
                if (i == j)
                {
                    phiSum += yj;
                }
                else
                {
                    phiSum += yj * WilkePhi(mu[i], mu[j], M[i], M[j]);
                }
            }

            muMix += yi * mu[i] / phiSum;

            // Wassiljewa-Mason-Saxena mixing rule for conductivity (same form as Wilke)
            double phiSumK = 0.0;
            for (int j = 0; j < n; j++)
            {
                double yj = molarFractions[j];
                if (yj <= 0.0) continue;
                if (i == j)
                {
                    phiSumK += yj;
                }
                else
                {
                    phiSumK += yj * WilkePhi(mu[i], mu[j], M[i], M[j]);
                }
            }
            kMix += yi * k[i] / phiSumK;
        }

        state.Viscosity = muMix;
        state.Conductivity = kMix;
        state.Prandtl = muMix * state.CpMass / kMix;
    }

    #endregion

    #region State solver (Newton-Raphson)

    /// <summary>
    /// Find temperature T given enthalpy H (J/kg) and pressure P (Pa).
    /// Uses Newton-Raphson with analytical derivative dH/dT = Cp.
    /// </summary>
    public static double SolveT_HP(double H_target, double P, int[] speciesIndices, double[] molarFractions, double Mmix)
    {
        // Initial guess: use T = 1000K as midpoint
        double T = 1000.0;
        const int maxIter = 50;
        const double tol = 1e-10;
        const double T_min = 200.0;
        const double T_max = 6000.0;

        for (int iter = 0; iter < maxIter; iter++)
        {
            // Compute H and Cp at current T
            double hMolar = 0.0, cpMolar = 0.0;
            for (int i = 0; i < speciesIndices.Length; i++)
            {
                double yi = molarFractions[i];
                if (yi <= 0.0) continue;
                int idx = speciesIndices[i];
                hMolar += yi * HOverRT(ref GasSpeciesData.Species[idx], T) * R * T;
                cpMolar += yi * CpOverR(ref GasSpeciesData.Species[idx], T) * R;
            }

            double hMass = hMolar / Mmix;
            double cpMass = cpMolar / Mmix;

            double dH = H_target - hMass;
            if (Math.Abs(dH) < tol * Math.Max(1.0, Math.Abs(H_target)))
                return T;

            // Newton step: T_new = T + (H_target - H(T)) / Cp(T)
            double dT = dH / cpMass;

            // Dampen and clamp
            if (T + dT < T_min) dT = (T_min - T) * 0.5;
            if (T + dT > T_max) dT = (T_max - T) * 0.5;

            T += dT;

            // Hard clamp
            if (T < T_min) T = T_min;
            if (T > T_max) T = T_max;
        }

        return T; // Return best estimate
    }

    /// <summary>
    /// Find temperature T given entropy S (J/(kg·K)) and pressure P (Pa).
    /// Uses Newton-Raphson with analytical derivative dS/dT = Cp/T.
    /// Entropy depends on pressure via S = S°(T) - R*ln(P/P_ref) / Mmix
    /// </summary>
    public static double SolveT_SP(double S_target, double P, int[] speciesIndices, double[] molarFractions, double Mmix)
    {
        double T = 1000.0;
        const int maxIter = 50;
        const double tol = 1e-10;
        const double T_min = 200.0;
        const double T_max = 6000.0;
        // Reference pressure for NASA entropy (1 atm = 101325 Pa)
        const double P_ref = 101325.0;
        double lnP_ratio = Math.Log(P / P_ref);

        for (int iter = 0; iter < maxIter; iter++)
        {
            double sMolar = 0.0, cpMolar = 0.0;
            for (int i = 0; i < speciesIndices.Length; i++)
            {
                double yi = molarFractions[i];
                if (yi <= 0.0) continue;
                int idx = speciesIndices[i];
                // S includes entropy of mixing: -R*yi*ln(yi)
                sMolar += yi * SOverR(ref GasSpeciesData.Species[idx], T) * R - R * yi * Math.Log(yi);
                cpMolar += yi * CpOverR(ref GasSpeciesData.Species[idx], T) * R;
            }
            // Pressure correction: S(T,P) = S°(T) - R*ln(P/P_ref)
            sMolar -= R * lnP_ratio;

            double sMass = sMolar / Mmix;
            double cpMass = cpMolar / Mmix;

            double dS = S_target - sMass;
            if (Math.Abs(dS) < tol * Math.Max(1.0, Math.Abs(S_target)))
                return T;

            // dS/dT = Cp/T → dT = dS * T / Cp
            double dT = dS * T / cpMass;

            if (T + dT < T_min) dT = (T_min - T) * 0.5;
            if (T + dT > T_max) dT = (T_max - T) * 0.5;

            T += dT;

            if (T < T_min) T = T_min;
            if (T > T_max) T = T_max;
        }

        return T;
    }

    #endregion

    #region High-level property resolver

    /// <summary>
    /// Resolve a single output property from a computed MixtureState.
    /// Returns NaN if property name is not recognized.
    /// Property names follow CoolProp conventions (after FormatName mapping).
    /// </summary>
    public static double GetProperty(ref MixtureState state, string property)
    {
        switch (property)
        {
            // Temperature
            case "T": return state.T;
            // Pressure
            case "P": return state.P;
            // Mass-specific properties
            case "Cpmass": case "CP": case "C": return state.CpMass;
            case "Cvmass": case "CV": return state.CvMass;
            case "H": return state.HMass;
            case "S": return state.SMass;
            case "U": return state.UMass;
            case "D": return state.Density;
            // Molar properties
            case "Cpmolar": return state.CpMolar;
            case "Cvmolar": return state.CvMolar;
            case "Hmolar": return state.HMolar;
            case "Smolar": return state.SMolar;
            // Derived
            case "isentropic_expansion_coefficient": return state.Gamma;
            case "A": return state.SpeedOfSound;
            case "M": return state.MolarMass;
            case "Z": return state.Z;
            // Transport
            case "viscosity": return state.Viscosity;
            case "conductivity": return state.Conductivity;
            case "Prandtl": return state.Prandtl;
            default: return double.NaN;
        }
    }

    /// <summary>
    /// Full property calculation: resolve state from input pair, compute all properties,
    /// return the requested output property.
    /// </summary>
    public static double CalcProperty(
        string output, string name1, double value1, string name2, double value2,
        int[] speciesIndices, double[] molarFractions)
    {
        // Pre-compute mixture molar mass
        double Mmix = 0.0;
        for (int i = 0; i < speciesIndices.Length; i++)
            Mmix += molarFractions[i] * GasSpeciesData.Species[speciesIndices[i]].MolarMass;

        double T, P;

        // Determine T and P from input pair
        string pair = name1 + "_" + name2;
        switch (pair)
        {
            case "T_P":
                T = value1; P = value2;
                break;
            case "P_T":
                T = value2; P = value1;
                break;
            case "H_P":
                P = value2;
                T = SolveT_HP(value1, P, speciesIndices, molarFractions, Mmix);
                break;
            case "P_H":
                P = value1;
                T = SolveT_HP(value2, P, speciesIndices, molarFractions, Mmix);
                break;
            case "S_P":
                P = value2;
                T = SolveT_SP(value1, P, speciesIndices, molarFractions, Mmix);
                break;
            case "P_S":
                P = value1;
                T = SolveT_SP(value2, P, speciesIndices, molarFractions, Mmix);
                break;
            default:
                return double.NaN;
        }

        // Validate temperature range
        if (T < 200.0 || T > 6000.0)
            return double.NaN;

        var state = CalcProperties(T, P, speciesIndices, molarFractions);
        return GetProperty(ref state, output);
    }

    #endregion
}
