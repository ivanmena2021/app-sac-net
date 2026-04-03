namespace Application.Services;

using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Application.DTOs.Responses;
using Domain.Entities;

/// <summary>Motor de consultas en lenguaje natural sobre datos SAC.</summary>
public class QueryEngineService
{
    private static readonly string[] Departamentos =
    {
        "AMAZONAS", "ANCASH", "APURIMAC", "AREQUIPA", "AYACUCHO",
        "CAJAMARCA", "CUSCO", "HUANCAVELICA", "HUANUCO", "ICA",
        "JUNIN", "LA LIBERTAD", "LAMBAYEQUE", "LIMA", "LORETO",
        "MADRE DE DIOS", "MOQUEGUA", "PASCO", "PIURA", "PUNO",
        "SAN MARTIN", "TACNA", "TUMBES", "UCAYALI",
    };

    private static readonly Dictionary<string, string> DeptoAliases = new(StringComparer.OrdinalIgnoreCase)
    {
        ["APURIMAC"] = "APURIMAC", ["JUNIN"] = "JUNIN", ["HUANUCO"] = "HUANUCO",
        ["ANCASH"] = "ANCASH", ["SAN MARTIN"] = "SAN MARTIN",
        ["CHICLAYO"] = "LAMBAYEQUE", ["TRUJILLO"] = "LA LIBERTAD",
        ["CUZCO"] = "CUSCO", ["IQUITOS"] = "LORETO", ["PUCALLPA"] = "UCAYALI",
        ["MOYOBAMBA"] = "SAN MARTIN", ["HUANCAYO"] = "JUNIN",
        ["CHACHAPOYAS"] = "AMAZONAS", ["ABANCAY"] = "APURIMAC",
        ["HUARAZ"] = "ANCASH", ["CERRO DE PASCO"] = "PASCO",
        ["PUERTO MALDONADO"] = "MADRE DE DIOS",
        ["MORROPE"] = "LAMBAYEQUE", ["OYOTUN"] = "LAMBAYEQUE",
        ["BOLIVAR"] = "LA LIBERTAD", ["NANCHOC"] = "CAJAMARCA",
    };

    private static readonly Dictionary<string, string[]> ConceptosSiniestro = new(StringComparer.OrdinalIgnoreCase)
    {
        ["lluvias"] = new[] { "INUNDACION", "LLUVIAS EXCESIVAS", "HUAYCO", "DESLIZAMIENTO" },
        ["lluvia"] = new[] { "INUNDACION", "LLUVIAS EXCESIVAS", "HUAYCO", "DESLIZAMIENTO" },
        ["exceso de agua"] = new[] { "INUNDACION", "LLUVIAS EXCESIVAS", "HUAYCO", "DESLIZAMIENTO" },
        ["frio"] = new[] { "HELADA", "FRIAJE", "NIEVE" },
        ["bajas temperaturas"] = new[] { "HELADA", "FRIAJE", "NIEVE" },
        ["climaticos"] = new[] { "HELADA", "SEQUIA", "GRANIZO", "INUNDACION", "LLUVIAS EXCESIVAS", "HUAYCO", "DESLIZAMIENTO", "VIENTO FUERTE", "NIEVE", "FRIAJE", "ALTAS TEMPERATURAS", "INCENDIO" },
        ["biologicos"] = new[] { "ENFERMEDADES", "PLAGAS" },
        ["fitosanitarios"] = new[] { "ENFERMEDADES", "PLAGAS" },
        ["sequia"] = new[] { "SEQUIA" },
        ["deficit hidrico"] = new[] { "SEQUIA" },
        ["calor"] = new[] { "ALTAS TEMPERATURAS", "INCENDIO" },
        ["vientos"] = new[] { "VIENTO FUERTE" },
    };

    private static readonly string[] TiposSiniestro =
    {
        "HELADA", "SEQUIA", "GRANIZO", "INUNDACION", "DESLIZAMIENTO",
        "ENFERMEDADES", "HUAYCO", "PLAGAS", "LLUVIAS EXCESIVAS",
        "VIENTO FUERTE", "INCENDIO", "ALTAS TEMPERATURAS", "NIEVE", "FRIAJE",
    };

    private static readonly Dictionary<string, string[]> MetricKeywords = new(StringComparer.OrdinalIgnoreCase)
    {
        ["avisos"] = new[] { "aviso", "avisos", "siniestro", "siniestros", "reportado", "reporte" },
        ["indemnizacion"] = new[] { "indemnizacion", "indemnizaciones", "indemnizado", "monto" },
        ["desembolso"] = new[] { "desembolso", "desembolsos", "desembolsado", "pagado", "pago" },
        ["productores"] = new[] { "productor", "productores", "beneficiado", "agricultor" },
        ["superficie"] = new[] { "hectarea", "hectareas", "superficie", "ha" },
        ["siniestralidad"] = new[] { "siniestralidad", "indice" },
        ["resumen"] = new[] { "resumen", "consolidado", "general" },
    };

    private static readonly Dictionary<string, int> Meses = new(StringComparer.OrdinalIgnoreCase)
    {
        ["enero"] = 1, ["febrero"] = 2, ["marzo"] = 3, ["abril"] = 4,
        ["mayo"] = 5, ["junio"] = 6, ["julio"] = 7, ["agosto"] = 8,
        ["septiembre"] = 9, ["octubre"] = 10, ["noviembre"] = 11, ["diciembre"] = 12,
    };

    private static string Normalize(string text)
    {
        var normalized = text.Normalize(NormalizationForm.FormKD);
        var sb = new StringBuilder();
        foreach (var c in normalized)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                sb.Append(c);
        }
        return sb.ToString().ToUpper().Trim();
    }

    public string ProcessQuery(string query, DatosNacionalesDto datos)
    {
        var midagri = datos.Midagri;
        var fechaCorte = datos.FechaCorte;

        // Detectar parametros
        var deptos = DetectDepartamentos(query);
        var tipos = DetectTiposSiniestro(query);
        var metrics = DetectMetrics(query);
        var temporal = DetectTemporal(query);
        var geoLevel = DetectGeographicLevel(query);

        // Filtrar datos
        var filtered = new List<Siniestro>(midagri);

        if (deptos.Count > 0)
            filtered = filtered.Where(s => deptos.Contains(s.Departamento.ToUpper())).ToList();

        if (tipos.Count > 0)
        {
            var tiposSet = new HashSet<string>(tipos, StringComparer.OrdinalIgnoreCase);
            filtered = filtered.Where(s => tiposSet.Contains(Normalize(s.TipoSiniestro ?? ""))).ToList();
        }

        // Filtrar por provincia/distrito (detectar de datos)
        var provincias = DetectFromData(query, midagri.Select(s => s.Provincia).Distinct());
        if (provincias.Count > 0)
            filtered = filtered.Where(s => provincias.Contains(s.Provincia?.ToUpper() ?? "")).ToList();

        var distritos = DetectFromData(query, midagri.Select(s => s.Distrito).Distinct());
        if (distritos.Count > 0)
            filtered = filtered.Where(s => distritos.Contains(s.Distrito?.ToUpper() ?? "")).ToList();

        // Filtrar temporal
        string? temporalLabel = null;
        if (temporal != null)
        {
            (filtered, temporalLabel) = ApplyTemporalFilter(filtered, temporal);
        }

        if (filtered.Count == 0)
        {
            var filters = new List<string>();
            if (deptos.Count > 0) filters.Add($"departamentos: {string.Join(", ", deptos.Select(ToTitle))}");
            if (tipos.Count > 0) filters.Add($"siniestros: {string.Join(", ", tipos.Select(ToTitle))}");
            if (temporalLabel != null) filters.Add($"periodo: {temporalLabel}");
            return $"No se encontraron registros con los filtros aplicados: {(filters.Count > 0 ? string.Join(", ", filters) : "ninguno")}.\n\nDatos disponibles al {fechaCorte}.";
        }

        var sb = new StringBuilder();

        // Header de contexto
        var contextParts = new List<string>();
        if (deptos.Count > 0) contextParts.Add($"**Departamentos:** {string.Join(", ", deptos.Select(ToTitle))}");
        if (tipos.Count > 0) contextParts.Add($"**Siniestros:** {string.Join(", ", tipos.Select(ToTitle))}");
        if (temporalLabel != null) contextParts.Add($"**Periodo:** {temporalLabel}");
        if (contextParts.Count > 0)
        {
            sb.AppendLine(string.Join(" · ", contextParts));
            sb.AppendLine("\n---\n");
        }

        // Agrupacion geografica
        if (geoLevel != null)
        {
            sb.Append(BuildGeographicSummary(filtered, geoLevel));
        }
        else if (deptos.Count > 0)
        {
            foreach (var depto in deptos)
            {
                var dfD = filtered.Where(s => s.Departamento.ToUpper() == depto).ToList();
                sb.AppendLine(BuildDeptoSummary(dfD, depto));
                sb.AppendLine();
            }
        }
        else
        {
            // Resumen general
            sb.AppendLine($"## Resumen SAC 2025-2026");
            sb.AppendLine($"**Fecha de corte:** {fechaCorte}\n");
            var totalIndemn = filtered.Sum(s => s.Indemnizacion);
            var totalDesemb = filtered.Sum(s => s.MontoDesembolsado);
            var totalProd = filtered.Sum(s => s.NProductores);
            var totalSup = filtered.Sum(s => s.SupIndemnizada);
            sb.AppendLine($"- **Avisos totales:** {filtered.Count:N0}");
            sb.AppendLine($"- **Indemnizacion reconocida:** S/ {totalIndemn:N2}");
            sb.AppendLine($"- **Desembolso:** S/ {totalDesemb:N2}");
            sb.AppendLine($"- **Productores:** {totalProd:N0}");
            if (totalSup > 0) sb.AppendLine($"- **Superficie indemnizada:** {totalSup:N2} ha");

            // Top departamentos
            var topDeptos = filtered.GroupBy(s => s.Departamento)
                .OrderByDescending(g => g.Count()).Take(5);
            sb.AppendLine("\n**Principales departamentos:**");
            foreach (var g in topDeptos)
                sb.AppendLine($"- {ToTitle(g.Key)}: {g.Count():N0} avisos");

            // Top tipos siniestro
            var topTipos = filtered.Where(s => !string.IsNullOrEmpty(s.TipoSiniestro))
                .GroupBy(s => s.TipoSiniestro).OrderByDescending(g => g.Count()).Take(5);
            sb.AppendLine("\n**Principales tipos de siniestro:**");
            foreach (var g in topTipos)
                sb.AppendLine($"- {ToTitle(g.Key)}: {g.Count():N0} avisos");
        }

        sb.AppendLine($"\n---\n*Fuente: DSFFA - MIDAGRI, SAC 2025-2026, datos al {fechaCorte}*");
        return sb.ToString();
    }

    private List<string> DetectDepartamentos(string query)
    {
        var queryNorm = Normalize(query);
        var found = new HashSet<string>();
        foreach (var (alias, depto) in DeptoAliases)
            if (queryNorm.Contains(Normalize(alias))) found.Add(depto);
        foreach (var depto in Departamentos)
            if (queryNorm.Contains(Normalize(depto))) found.Add(depto);
        return found.OrderBy(d => d).ToList();
    }

    private List<string> DetectTiposSiniestro(string query)
    {
        var queryNorm = Normalize(query);
        var queryLower = query.ToLower();
        var found = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var (concepto, tipos) in ConceptosSiniestro)
            if (queryLower.Contains(concepto))
                foreach (var t in tipos) found.Add(t);

        foreach (var tipo in TiposSiniestro)
            if (queryNorm.Contains(Normalize(tipo))) found.Add(tipo);

        return found.OrderBy(t => t).ToList();
    }

    private HashSet<string> DetectMetrics(string query)
    {
        var queryLower = query.ToLower();
        var found = new HashSet<string>();
        foreach (var (metric, keywords) in MetricKeywords)
            foreach (var kw in keywords)
                if (queryLower.Contains(kw)) { found.Add(metric); break; }
        if (found.Count == 0) found.Add("resumen");
        return found;
    }

    private Dictionary<string, object>? DetectTemporal(string query)
    {
        var queryLower = query.ToLower();
        var yearMatch = Regex.Match(queryLower, @"\b(20[2-3]\d)\b");
        int? year = yearMatch.Success ? int.Parse(yearMatch.Groups[1].Value) : null;
        int? month = null;
        foreach (var (name, num) in Meses)
            if (queryLower.Contains(name)) { month = num; break; }

        if (year.HasValue && month.HasValue)
            return new Dictionary<string, object> { ["type"] = "year_month", ["year"] = year.Value, ["month"] = month.Value };
        if (year.HasValue)
            return new Dictionary<string, object> { ["type"] = "year", ["year"] = year.Value };
        if (month.HasValue)
            return new Dictionary<string, object> { ["type"] = "month", ["month"] = month.Value };

        var temporalKw = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["ultima semana"] = 7, ["semana"] = 7, ["mes"] = 30, ["ultimo mes"] = 30,
            ["quincena"] = 15, ["hoy"] = 1, ["ayer"] = 2,
        };
        foreach (var (kw, days) in temporalKw.OrderByDescending(k => k.Key.Length))
            if (queryLower.Contains(kw))
                return new Dictionary<string, object> { ["type"] = "days", ["days"] = days };

        return null;
    }

    private string? DetectGeographicLevel(string query)
    {
        var q = query.ToLower();
        if (q.Contains("por distrito") || q.Contains("nivel distrital")) return "distrito";
        if (q.Contains("por provincia") || q.Contains("nivel provincial")) return "provincia";
        return null;
    }

    private List<string> DetectFromData(string query, IEnumerable<string?> values)
    {
        var queryNorm = Normalize(query);
        var found = new HashSet<string>();
        foreach (var val in values)
        {
            if (string.IsNullOrWhiteSpace(val) || val.Length < 4) continue;
            var clean = val.Trim().ToUpper();
            if (Normalize(clean).Length >= 4 && queryNorm.Contains(Normalize(clean)))
                found.Add(clean);
        }
        return found.ToList();
    }

    private (List<Siniestro>, string?) ApplyTemporalFilter(List<Siniestro> data, Dictionary<string, object> temporal)
    {
        var type = temporal["type"].ToString();
        string? label = null;

        if (type == "days")
        {
            var days = Convert.ToInt32(temporal["days"]);
            var cutoff = DateTime.Now.AddDays(-days);
            data = data.Where(s => (s.FechaAviso ?? s.FechaSiniestro) >= cutoff).ToList();
            label = $"ultimos {days} dias";
        }
        else if (type == "year")
        {
            var yr = Convert.ToInt32(temporal["year"]);
            data = data.Where(s => (s.FechaAviso ?? s.FechaSiniestro)?.Year == yr).ToList();
            label = $"ano {yr}";
        }
        else if (type == "year_month")
        {
            var yr = Convert.ToInt32(temporal["year"]);
            var mo = Convert.ToInt32(temporal["month"]);
            data = data.Where(s =>
            {
                var f = s.FechaAviso ?? s.FechaSiniestro;
                return f?.Year == yr && f?.Month == mo;
            }).ToList();
            var mesName = Meses.FirstOrDefault(m => m.Value == mo).Key ?? mo.ToString();
            label = $"{ToTitle(mesName)} {yr}";
        }
        else if (type == "month")
        {
            var mo = Convert.ToInt32(temporal["month"]);
            data = data.Where(s => (s.FechaAviso ?? s.FechaSiniestro)?.Month == mo).ToList();
            var mesName = Meses.FirstOrDefault(m => m.Value == mo).Key ?? mo.ToString();
            label = ToTitle(mesName);
        }

        return (data, label);
    }

    private string BuildDeptoSummary(List<Siniestro> records, string depto)
    {
        if (records.Count == 0)
            return $"**{ToTitle(depto)}**: Sin avisos registrados.";

        var sb = new StringBuilder();
        sb.AppendLine($"### {ToTitle(depto)}");
        sb.AppendLine($"- **Avisos reportados:** {records.Count:N0}");

        var cerrados = records.Count(s => s.EstadoInspeccion?.ToUpper() == "CERRADO");
        var pctEval = records.Count > 0 ? (100.0 * cerrados / records.Count) : 0;
        sb.AppendLine($"- **Evaluados (cerrados):** {cerrados:N0}");
        sb.AppendLine($"- **Avance de evaluacion:** {pctEval:F1}%");

        var indemn = records.Sum(s => s.Indemnizacion);
        sb.AppendLine($"- **Indemnizacion reconocida:** S/ {indemn:N2}");

        var sup = records.Sum(s => s.SupIndemnizada);
        if (sup > 0) sb.AppendLine($"- **Superficie indemnizada:** {sup:N2} ha");

        var desemb = records.Sum(s => s.MontoDesembolsado);
        var pctD = indemn > 0 ? (100.0 * desemb / indemn) : 0;
        sb.AppendLine($"- **Desembolso:** S/ {desemb:N2}");
        sb.AppendLine($"- **Avance de desembolso:** {pctD:F1}%");

        var prod = records.Sum(s => s.NProductores);
        if (prod > 0) sb.AppendLine($"- **Productores beneficiados:** {prod:N0}");

        var topTipos = records.Where(s => !string.IsNullOrEmpty(s.TipoSiniestro))
            .GroupBy(s => s.TipoSiniestro).OrderByDescending(g => g.Count()).Take(5);
        if (topTipos.Any())
        {
            var tiposText = string.Join(", ", topTipos.Select(g => $"{ToTitle(g.Key)} ({g.Count():N0})"));
            sb.AppendLine($"- **Principales siniestros:** {tiposText}");
        }

        return sb.ToString();
    }

    private string BuildGeographicSummary(List<Siniestro> records, string level)
    {
        Func<Siniestro, string> groupBy = level == "provincia"
            ? s => s.Provincia?.Trim().ToUpper() ?? ""
            : s => s.Distrito?.Trim().ToUpper() ?? "";
        var label = level == "provincia" ? "Provincia" : "Distrito";

        var sb = new StringBuilder();
        sb.AppendLine($"## Resumen por {label}\n");
        sb.AppendLine($"**Total de avisos:** {records.Count:N0}\n");

        var grouped = records.GroupBy(groupBy)
            .Where(g => g.Key.Length > 0)
            .OrderByDescending(g => g.Count()).Take(15);

        foreach (var g in grouped)
        {
            var n = g.Count();
            var pct = records.Count > 0 ? (100.0 * n / records.Count) : 0;
            sb.AppendLine($"### {ToTitle(g.Key)} ({n:N0} avisos — {pct:F1}%)");
            var gIndemn = g.Sum(s => s.Indemnizacion);
            var gDesemb = g.Sum(s => s.MontoDesembolsado);
            if (gIndemn > 0) sb.Append($"- Indemnizacion: S/ {gIndemn:N2}");
            if (gDesemb > 0) sb.Append($" | Desembolso: S/ {gDesemb:N2}");
            sb.AppendLine();
        }

        return sb.ToString();
    }

    private static string ToTitle(string s) =>
        CultureInfo.CurrentCulture.TextInfo.ToTitleCase(s.ToLower());

    public static List<string> GetSuggestedQueries() => new()
    {
        "Resumen de Tumbes, Piura, Lambayeque, Lima y Arequipa",
        "Intervenciones del SAC en Cajamarca y Lambayeque",
        "Cuantos avisos tiene Ayacucho?",
        "Desembolsos en Junin y Cusco",
        "Avisos por eventos asociados a lluvias en 2026",
        "Heladas y frio en Puno y Huancavelica",
        "Avisos por provincia en Cusco",
        "Resumen por distrito en Lambayeque",
    };
}
