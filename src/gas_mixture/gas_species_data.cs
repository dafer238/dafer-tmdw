using System;
using System.Collections.Generic;

/// <summary>
/// Static data for ideal gas species: NASA 7-coefficient polynomials,
/// molar masses, and Sutherland transport parameters.
/// NASA coefficients sourced from GRI-Mech 3.0 / Burcat thermodynamic database.
/// </summary>
public static class GasSpeciesData
{
    /// <summary>
    /// NASA 7-coefficient polynomial data for a single temperature range.
    /// Cp/R = a1 + a2*T + a3*T^2 + a4*T^3 + a5*T^4
    /// H/(RT) = a1 + a2*T/2 + a3*T^2/3 + a4*T^3/4 + a5*T^4/5 + a6/T
    /// S/R = a1*ln(T) + a2*T + a3*T^2/2 + a4*T^3/3 + a5*T^4/4 + a7
    /// </summary>
    public struct NasaCoeffs
    {
        public double A1, A2, A3, A4, A5, A6, A7;

        public NasaCoeffs(double a1, double a2, double a3, double a4, double a5, double a6, double a7)
        {
            A1 = a1; A2 = a2; A3 = a3; A4 = a4; A5 = a5; A6 = a6; A7 = a7;
        }
    }

    /// <summary>Sutherland's law parameters for viscosity or conductivity.</summary>
    public struct SutherlandParams
    {
        public double RefValue; // μ_ref or k_ref at T_ref
        public double TRef;     // Reference temperature (K)
        public double S;        // Sutherland constant (K)

        public SutherlandParams(double refValue, double tRef, double s)
        {
            RefValue = refValue; TRef = tRef; S = s;
        }
    }

    /// <summary>Complete data for a single gas species.</summary>
    public struct GasSpecies
    {
        public string Name;        // Canonical name (e.g. "N2")
        public double MolarMass;   // kg/mol
        public double TLow;        // Low T range start (K)
        public double TMid;        // Switchover temperature (K)
        public double THigh;       // High T range end (K)
        public NasaCoeffs LowCoeffs;   // Valid for TLow <= T <= TMid
        public NasaCoeffs HighCoeffs;  // Valid for TMid < T <= THigh
        public SutherlandParams ViscParams;  // Sutherland viscosity parameters
        public SutherlandParams CondParams;  // Sutherland conductivity parameters
    }

    // Universal gas constant (J/(mol·K))
    public const double R = 8.314462618;

    /// <summary>All species data, indexed by species ID.</summary>
    public static readonly GasSpecies[] Species;

    /// <summary>Maps gas name aliases (case-insensitive) to species index in Species array.</summary>
    public static readonly Dictionary<string, int> NameToIndex;

    static GasSpeciesData()
    {
        // Define all species
        Species = new GasSpecies[]
        {
            // 0: N2 - Nitrogen (GRI-Mech 3.0)
            new GasSpecies
            {
                Name = "N2", MolarMass = 0.0280134,
                TLow = 200.0, TMid = 1000.0, THigh = 6000.0,
                LowCoeffs = new NasaCoeffs(
                    3.53100528E+00, -1.23660988E-04, -5.02999433E-07,
                    2.43530612E-09, -1.40881235E-12, -1.04697628E+03, 2.96747038E+00),
                HighCoeffs = new NasaCoeffs(
                    2.95257637E+00,  1.39690040E-03, -4.92631603E-07,
                    7.86010195E-11, -4.60755204E-15, -9.23948688E+02, 5.87188762E+00),
                ViscParams = new SutherlandParams(1.663e-5, 273.15, 107.0),
                CondParams = new SutherlandParams(2.40e-2, 273.15, 150.0)
            },
            // 1: O2 - Oxygen (GRI-Mech 3.0)
            new GasSpecies
            {
                Name = "O2", MolarMass = 0.0319988,
                TLow = 200.0, TMid = 1000.0, THigh = 6000.0,
                LowCoeffs = new NasaCoeffs(
                    3.78245636E+00, -2.99673416E-03,  9.84730201E-06,
                   -9.68129509E-09,  3.24372837E-12, -1.06394356E+03, 3.65767573E+00),
                HighCoeffs = new NasaCoeffs(
                    3.66096065E+00,  6.56365811E-04, -1.41149627E-07,
                    2.05797935E-11, -1.29913436E-15, -1.21597718E+03, 3.41536279E+00),
                ViscParams = new SutherlandParams(1.919e-5, 273.15, 139.0),
                CondParams = new SutherlandParams(2.44e-2, 273.15, 240.0)
            },
            // 2: CO2 - Carbon Dioxide (GRI-Mech 3.0)
            new GasSpecies
            {
                Name = "CO2", MolarMass = 0.04400995,
                TLow = 200.0, TMid = 1000.0, THigh = 6000.0,
                LowCoeffs = new NasaCoeffs(
                    2.35677352E+00,  8.98459677E-03, -7.12356269E-06,
                    2.45919022E-09, -1.43699548E-13, -4.83719697E+04, 9.90105222E+00),
                HighCoeffs = new NasaCoeffs(
                    4.63659493E+00,  2.74131991E-03, -9.95828531E-07,
                    1.60373011E-10, -9.16103468E-15, -4.90249341E+04, -1.93534855E+00),
                ViscParams = new SutherlandParams(1.370e-5, 273.15, 222.0),
                CondParams = new SutherlandParams(1.45e-2, 273.15, 1800.0)
            },
            // 3: H2O - Water Vapor (GRI-Mech 3.0)
            new GasSpecies
            {
                Name = "H2O", MolarMass = 0.01801528,
                TLow = 200.0, TMid = 1000.0, THigh = 6000.0,
                LowCoeffs = new NasaCoeffs(
                    4.19864056E+00, -2.03643410E-03,  6.52040211E-06,
                   -5.48797062E-09,  1.77197817E-12, -3.02937267E+04, -8.49032208E-01),
                HighCoeffs = new NasaCoeffs(
                    2.67703787E+00,  2.97318160E-03, -7.73769690E-07,
                    9.44336689E-11, -4.26900959E-15, -2.98858938E+04, 6.88255571E+00),
                ViscParams = new SutherlandParams(1.12e-5, 350.0, 1064.0),
                CondParams = new SutherlandParams(1.81e-2, 350.0, 2200.0)
            },
            // 4: Ar - Argon (GRI-Mech 3.0)
            new GasSpecies
            {
                Name = "Ar", MolarMass = 0.039948,
                TLow = 200.0, TMid = 1000.0, THigh = 6000.0,
                LowCoeffs = new NasaCoeffs(
                    2.50000000E+00,  0.00000000E+00,  0.00000000E+00,
                    0.00000000E+00,  0.00000000E+00, -7.45375000E+02, 4.37967491E+00),
                HighCoeffs = new NasaCoeffs(
                    2.50000000E+00,  0.00000000E+00,  0.00000000E+00,
                    0.00000000E+00,  0.00000000E+00, -7.45375000E+02, 4.37967491E+00),
                ViscParams = new SutherlandParams(2.125e-5, 273.15, 114.0),
                CondParams = new SutherlandParams(1.63e-2, 273.15, 170.0)
            },
            // 5: CO - Carbon Monoxide (GRI-Mech 3.0)
            new GasSpecies
            {
                Name = "CO", MolarMass = 0.0280101,
                TLow = 200.0, TMid = 1000.0, THigh = 6000.0,
                LowCoeffs = new NasaCoeffs(
                    3.57953347E+00, -6.10353680E-04,  1.01681433E-06,
                    9.07005884E-10, -9.04424499E-13, -1.43440860E+04, 3.50840928E+00),
                HighCoeffs = new NasaCoeffs(
                    3.04848583E+00,  1.35172818E-03, -4.85794075E-07,
                    7.88536486E-11, -4.69807489E-15, -1.42677477E+04, 6.01709790E+00),
                ViscParams = new SutherlandParams(1.657e-5, 273.15, 136.0),
                CondParams = new SutherlandParams(2.32e-2, 273.15, 180.0)
            },
            // 6: H2 - Hydrogen (GRI-Mech 3.0)
            new GasSpecies
            {
                Name = "H2", MolarMass = 0.00201588,
                TLow = 200.0, TMid = 1000.0, THigh = 6000.0,
                LowCoeffs = new NasaCoeffs(
                    2.34433112E+00,  7.98052075E-03, -1.94781510E-05,
                    2.01572094E-08, -7.37611761E-12, -9.17935173E+02, 6.83010238E-01),
                HighCoeffs = new NasaCoeffs(
                    2.93286575E+00,  8.26608026E-04, -1.46402364E-07,
                    1.54100414E-11, -6.88804800E-16, -8.13065581E+02, -1.02432865E+00),
                ViscParams = new SutherlandParams(8.411e-6, 273.15, 97.0),
                CondParams = new SutherlandParams(1.68e-1, 273.15, 120.0)
            },
            // 7: NO - Nitric Oxide (Burcat)
            new GasSpecies
            {
                Name = "NO", MolarMass = 0.0300061,
                TLow = 200.0, TMid = 1000.0, THigh = 6000.0,
                LowCoeffs = new NasaCoeffs(
                    4.21859896E+00, -4.63988124E-03,  1.10443049E-05,
                   -9.34055507E-09,  2.80554874E-12,  9.84509964E+03, 2.28061001E+00),
                HighCoeffs = new NasaCoeffs(
                    3.26071234E+00,  1.19101135E-03, -4.29122646E-07,
                    6.94481463E-11, -4.03295681E-15,  9.92143132E+03, 6.36900518E+00),
                ViscParams = new SutherlandParams(1.78e-5, 273.15, 128.0),
                CondParams = new SutherlandParams(2.36e-2, 273.15, 150.0)
            },
            // 8: NO2 - Nitrogen Dioxide (Burcat)
            new GasSpecies
            {
                Name = "NO2", MolarMass = 0.0460055,
                TLow = 200.0, TMid = 1000.0, THigh = 6000.0,
                LowCoeffs = new NasaCoeffs(
                    3.94403120E+00, -1.58542900E-03,  1.66578120E-05,
                   -2.04754260E-08,  7.83505640E-12,  2.89661800E+03, 6.31199190E+00),
                HighCoeffs = new NasaCoeffs(
                    4.88475400E+00,  2.17239570E-03, -8.28069090E-07,
                    1.57475100E-10, -1.05108950E-14,  2.31649800E+03, -1.17416951E-01),
                ViscParams = new SutherlandParams(1.37e-5, 273.15, 220.0),
                CondParams = new SutherlandParams(1.45e-2, 273.15, 500.0)
            },
            // 9: CH4 - Methane (GRI-Mech 3.0)
            new GasSpecies
            {
                Name = "CH4", MolarMass = 0.01604246,
                TLow = 200.0, TMid = 1000.0, THigh = 6000.0,
                LowCoeffs = new NasaCoeffs(
                    5.14987613E+00, -1.36709788E-02,  4.91800599E-05,
                   -4.84743026E-08,  1.66693956E-11, -1.02466476E+04, -4.64130376E+00),
                HighCoeffs = new NasaCoeffs(
                    1.65326226E+00,  1.00263099E-02, -3.31661238E-06,
                    5.36483138E-10, -3.14696758E-14, -1.00095936E+04, 9.90506283E+00),
                ViscParams = new SutherlandParams(1.024e-5, 273.15, 164.0),
                CondParams = new SutherlandParams(3.02e-2, 273.15, 164.0)
            },
            // 10: N2O - Nitrous Oxide (Burcat)
            new GasSpecies
            {
                Name = "N2O", MolarMass = 0.0440128,
                TLow = 200.0, TMid = 1000.0, THigh = 6000.0,
                LowCoeffs = new NasaCoeffs(
                    2.25715020E+00,  1.13047280E-02, -1.36713190E-05,
                    9.68198030E-09, -2.93071820E-12,  8.74177400E+03, 1.07579920E+01),
                HighCoeffs = new NasaCoeffs(
                    4.82307290E+00,  2.62702510E-03, -9.58508720E-07,
                    1.60007090E-10, -9.77523050E-15,  8.07340480E+03, -2.20172080E+00),
                ViscParams = new SutherlandParams(1.37e-5, 273.15, 240.0),
                CondParams = new SutherlandParams(1.51e-2, 273.15, 500.0)
            },
            // 11: SO2 - Sulfur Dioxide (Burcat)
            new GasSpecies
            {
                Name = "SO2", MolarMass = 0.0640638,
                TLow = 200.0, TMid = 1000.0, THigh = 6000.0,
                LowCoeffs = new NasaCoeffs(
                    3.26653380E+00,  5.32379020E-03,  6.84375520E-07,
                   -5.28100470E-09,  2.55904540E-12, -3.69081480E+04, 9.66465108E+00),
                HighCoeffs = new NasaCoeffs(
                    5.24513640E+00,  1.97042040E-03, -8.03757690E-07,
                    1.51499690E-10, -1.05580040E-14, -3.75582270E+04, -1.07404892E+00),
                ViscParams = new SutherlandParams(1.17e-5, 273.15, 306.0),
                CondParams = new SutherlandParams(8.40e-3, 273.15, 700.0)
            },
            // 12: He - Helium (monatomic ideal gas)
            new GasSpecies
            {
                Name = "He", MolarMass = 0.0040026,
                TLow = 200.0, TMid = 1000.0, THigh = 6000.0,
                LowCoeffs = new NasaCoeffs(
                    2.50000000E+00,  0.00000000E+00,  0.00000000E+00,
                    0.00000000E+00,  0.00000000E+00, -7.45375000E+02, 9.28723974E-01),
                HighCoeffs = new NasaCoeffs(
                    2.50000000E+00,  0.00000000E+00,  0.00000000E+00,
                    0.00000000E+00,  0.00000000E+00, -7.45375000E+02, 9.28723974E-01),
                ViscParams = new SutherlandParams(1.864e-5, 273.15, 79.4),
                CondParams = new SutherlandParams(1.42e-1, 273.15, 79.4)
            },
            // 13: Ne - Neon (monatomic ideal gas)
            new GasSpecies
            {
                Name = "Ne", MolarMass = 0.0201797,
                TLow = 200.0, TMid = 1000.0, THigh = 6000.0,
                LowCoeffs = new NasaCoeffs(
                    2.50000000E+00,  0.00000000E+00,  0.00000000E+00,
                    0.00000000E+00,  0.00000000E+00, -7.45375000E+02, 3.35532272E+00),
                HighCoeffs = new NasaCoeffs(
                    2.50000000E+00,  0.00000000E+00,  0.00000000E+00,
                    0.00000000E+00,  0.00000000E+00, -7.45375000E+02, 3.35532272E+00),
                ViscParams = new SutherlandParams(2.985e-5, 273.15, 56.0),
                CondParams = new SutherlandParams(4.61e-2, 273.15, 56.0)
            },
            // 14: NH3 - Ammonia (Burcat)
            new GasSpecies
            {
                Name = "NH3", MolarMass = 0.01703052,
                TLow = 200.0, TMid = 1000.0, THigh = 6000.0,
                LowCoeffs = new NasaCoeffs(
                    4.28648780E+00, -4.66056060E-03,  2.17159920E-05,
                   -2.28067060E-08,  8.26382810E-12, -6.74130070E+03, -6.25371230E-01),
                HighCoeffs = new NasaCoeffs(
                    2.63455280E+00,  5.66614900E-03, -1.72791830E-06,
                    2.38679830E-10, -1.23698050E-14, -6.54425630E+03, 6.56630940E+00),
                ViscParams = new SutherlandParams(9.35e-6, 273.15, 370.0),
                CondParams = new SutherlandParams(2.18e-2, 273.15, 370.0)
            },
            // 15: C2H6 - Ethane (GRI-Mech 3.0)
            new GasSpecies
            {
                Name = "C2H6", MolarMass = 0.03006904,
                TLow = 200.0, TMid = 1000.0, THigh = 6000.0,
                LowCoeffs = new NasaCoeffs(
                    4.29142492E+00, -5.50154270E-03,  5.99438288E-05,
                   -7.08466285E-08,  2.68685771E-11, -1.15222055E+04, 2.66682316E+00),
                HighCoeffs = new NasaCoeffs(
                    4.04666411E+00,  1.53538802E-02, -5.47039485E-06,
                    8.77826544E-10, -5.23167531E-14, -1.24473499E+04, -9.68698313E-01),
                ViscParams = new SutherlandParams(8.60e-6, 273.15, 252.0),
                CondParams = new SutherlandParams(1.83e-2, 273.15, 252.0)
            },
            // 16: C3H8 - Propane (Burcat)
            new GasSpecies
            {
                Name = "C3H8", MolarMass = 0.04409562,
                TLow = 200.0, TMid = 1000.0, THigh = 6000.0,
                LowCoeffs = new NasaCoeffs(
                    4.21093001E+00,  1.70886866E-03,  7.06529457E-05,
                   -9.20060869E-08,  3.64618834E-11, -1.43953975E+04, 5.61226917E+00),
                HighCoeffs = new NasaCoeffs(
                    6.66919718E+00,  2.06140622E-02, -7.36507688E-06,
                    1.18431786E-09, -7.06777994E-14, -1.62750939E+04, -1.31440251E+01),
                ViscParams = new SutherlandParams(7.50e-6, 273.15, 338.0),
                CondParams = new SutherlandParams(1.52e-2, 273.15, 338.0)
            },
            // 17: C4H10 - n-Butane (Burcat)
            new GasSpecies
            {
                Name = "C4H10", MolarMass = 0.05812220,
                TLow = 200.0, TMid = 1000.0, THigh = 6000.0,
                LowCoeffs = new NasaCoeffs(
                    5.14987750E+00, -3.41023780E-03,  1.10897250E-04,
                   -1.43826218E-07,  5.71234630E-11, -1.73879206E+04, 2.85498650E+00),
                HighCoeffs = new NasaCoeffs(
                    9.32553100E+00,  2.58498000E-02, -9.23798350E-06,
                    1.48755200E-09, -8.88250000E-14, -2.00075320E+04, -2.55828480E+01),
                ViscParams = new SutherlandParams(6.90e-6, 273.15, 377.0),
                CondParams = new SutherlandParams(1.35e-2, 273.15, 377.0)
            },
            // 18: iC4H10 - Isobutane (Burcat)
            new GasSpecies
            {
                Name = "iC4H10", MolarMass = 0.05812220,
                TLow = 200.0, TMid = 1000.0, THigh = 6000.0,
                LowCoeffs = new NasaCoeffs(
                    4.76013400E+00,  3.88609300E-04,  1.02850000E-04,
                   -1.35800000E-07,  5.48930000E-11, -1.80254000E+04, 3.72949200E+00),
                HighCoeffs = new NasaCoeffs(
                    9.12624000E+00,  2.60690000E-02, -9.30654000E-06,
                    1.49660000E-09, -8.93150000E-14, -2.05834000E+04, -2.46155000E+01),
                ViscParams = new SutherlandParams(6.80e-6, 273.15, 372.0),
                CondParams = new SutherlandParams(1.42e-2, 273.15, 372.0)
            },
            // 19: C5H12 - n-Pentane (Burcat)
            new GasSpecies
            {
                Name = "C5H12", MolarMass = 0.07214878,
                TLow = 200.0, TMid = 1000.0, THigh = 6000.0,
                LowCoeffs = new NasaCoeffs(
                    5.76346350E+00, -5.25266400E-04,  1.39224250E-04,
                   -1.81513950E-07,  7.24780000E-11, -2.01568000E+04, 2.20780000E+00),
                HighCoeffs = new NasaCoeffs(
                    1.18766400E+01,  3.12080000E-02, -1.11530000E-05,
                    1.79690000E-09, -1.07310000E-13, -2.32120000E+04, -3.78318000E+01),
                ViscParams = new SutherlandParams(6.30e-6, 273.15, 408.0),
                CondParams = new SutherlandParams(1.31e-2, 273.15, 408.0)
            },
            // 20: iC5H12 - Isopentane (Burcat)
            new GasSpecies
            {
                Name = "iC5H12", MolarMass = 0.07214878,
                TLow = 200.0, TMid = 1000.0, THigh = 6000.0,
                LowCoeffs = new NasaCoeffs(
                    5.54600000E+00,  1.05200000E-03,  1.28890000E-04,
                   -1.72200000E-07,  6.94300000E-11, -2.07080000E+04, 2.43200000E+00),
                HighCoeffs = new NasaCoeffs(
                    1.16150000E+01,  3.15300000E-02, -1.12610000E-05,
                    1.81400000E-09, -1.08270000E-13, -2.37490000E+04, -3.63400000E+01),
                ViscParams = new SutherlandParams(6.25e-6, 273.15, 400.0),
                CondParams = new SutherlandParams(1.26e-2, 273.15, 400.0)
            },
        };

        // Build name-to-index mapping (case-insensitive)
        NameToIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < Species.Length; i++)
        {
            NameToIndex[Species[i].Name] = i;
        }

        // Add common aliases
        NameToIndex["NITROGEN"] = 0;
        NameToIndex["OXYGEN"] = 1;
        NameToIndex["CARBONDIOXIDE"] = 2;
        NameToIndex["CARBON DIOXIDE"] = 2;
        NameToIndex["WATER"] = 3;
        NameToIndex["WATERVAPOR"] = 3;
        NameToIndex["WATER VAPOR"] = 3;
        NameToIndex["STEAM"] = 3;
        NameToIndex["ARGON"] = 4;
        NameToIndex["CARBONMONOXIDE"] = 5;
        NameToIndex["CARBON MONOXIDE"] = 5;
        NameToIndex["HYDROGEN"] = 6;
        NameToIndex["NITRICOXIDE"] = 7;
        NameToIndex["NITRIC OXIDE"] = 7;
        NameToIndex["NITROGENDIOXIDE"] = 8;
        NameToIndex["NITROGEN DIOXIDE"] = 8;
        NameToIndex["METHANE"] = 9;
        NameToIndex["NITROUSOXIDE"] = 10;
        NameToIndex["NITROUS OXIDE"] = 10;
        NameToIndex["SULFURDIOXIDE"] = 11;
        NameToIndex["SULFUR DIOXIDE"] = 11;
        NameToIndex["HELIUM"] = 12;
        NameToIndex["NEON"] = 13;
        NameToIndex["AMMONIA"] = 14;
        NameToIndex["ETHANE"] = 15;
        NameToIndex["PROPANE"] = 16;
        NameToIndex["BUTANE"] = 17;
        NameToIndex["NBUTANE"] = 17;
        NameToIndex["N-BUTANE"] = 17;
        NameToIndex["ISOBUTANE"] = 18;
        NameToIndex["IBUTANE"] = 18;
        NameToIndex["I-BUTANE"] = 18;
        NameToIndex["IC4H10"] = 18;
        NameToIndex["PENTANE"] = 19;
        NameToIndex["NPENTANE"] = 19;
        NameToIndex["N-PENTANE"] = 19;
        NameToIndex["ISOPENTANE"] = 20;
        NameToIndex["IPENTANE"] = 20;
        NameToIndex["I-PENTANE"] = 20;
        NameToIndex["IC5H12"] = 20;
    }

    /// <summary>
    /// Resolve a gas name string to a species index. Returns -1 if not found.
    /// </summary>
    public static int ResolveSpecies(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return -1;
        string trimmed = name.Trim();
        if (NameToIndex.TryGetValue(trimmed, out int index))
            return index;
        // Try without dashes, underscores, spaces
        string normalized = trimmed.Replace("-", "").Replace("_", "").Replace(" ", "");
        if (NameToIndex.TryGetValue(normalized, out index))
            return index;
        return -1;
    }
}
