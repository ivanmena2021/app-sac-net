namespace Infrastructure.StaticData;

using System.Data;
using System.Globalization;
using System.Text;
using System.Text.Json;
using Application.Contracts.Repositories;
using Application.DTOs.Responses;
using Domain.Entities;
using ExcelDataReader;
using Microsoft.Extensions.Hosting;

public class StaticDataRepository : IStaticDataRepository
{
    private readonly string _dataPath;

    static StaticDataRepository()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public StaticDataRepository(IHostEnvironment env)
    {
        _dataPath = Path.Combine(env.ContentRootPath, "wwwroot", "data");
    }

    // ── Normalización de acentos ─────────────────────────────
    private static string NormalizeDepto(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return name;
        name = name.Trim().ToUpper();
        var map = new Dictionary<char, char>
        {
            ['\u00C1'] = 'A', ['\u00C9'] = 'E', ['\u00CD'] = 'I', ['\u00D3'] = 'O', ['\u00DA'] = 'U',
            ['\u00E1'] = 'A', ['\u00E9'] = 'E', ['\u00ED'] = 'I', ['\u00F3'] = 'O', ['\u00FA'] = 'U',
            ['\u00D1'] = 'N', ['\u00F1'] = 'N',
        };
        var sb = new StringBuilder(name.Length);
        foreach (var c in name)
            sb.Append(map.TryGetValue(c, out var r) ? r : c);
        return sb.ToString();
    }

    // ── Resumen departamental ────────────────────────────────
    public Dictionary<string, Dictionary<string, DeptoCampanaDto>> LoadResumenDepartamental()
    {
        var filePath = Path.Combine(_dataPath, "resumen_departamental.json");
        if (!File.Exists(filePath)) return new();

        var json = File.ReadAllText(filePath, Encoding.UTF8);
        using var doc = JsonDocument.Parse(json);
        var result = new Dictionary<string, Dictionary<string, DeptoCampanaDto>>(StringComparer.OrdinalIgnoreCase);

        if (!doc.RootElement.TryGetProperty("por_campana", out var porCampana))
            return result;

        foreach (var deptProp in porCampana.EnumerateObject())
        {
            var campanas = new Dictionary<string, DeptoCampanaDto>();
            foreach (var campProp in deptProp.Value.EnumerateObject())
            {
                var c = campProp.Value;
                campanas[campProp.Name] = new DeptoCampanaDto
                {
                    Avisos = c.TryGetProperty("avisos", out var a) ? a.GetInt32() : 0,
                    Indemnizados = c.TryGetProperty("indemnizados", out var i) ? i.GetInt32() : 0,
                    PerdidaTotal = c.TryGetProperty("perdida_total", out var pt) ? pt.GetInt32() : 0,
                    MontoIndemnizado = c.TryGetProperty("monto_indemnizado", out var mi) ? mi.GetDouble() : 0,
                    HaIndemnizadas = c.TryGetProperty("ha_indemnizadas", out var ha) ? ha.GetDouble() : 0,
                    MontoDesembolsado = c.TryGetProperty("monto_desembolsado", out var md) ? md.GetDouble() : 0,
                    Provincias = c.TryGetProperty("provincias", out var pv) ? pv.GetInt32() : 0,
                    Distritos = c.TryGetProperty("distritos", out var di) ? di.GetInt32() : 0,
                };
            }
            result[deptProp.Name] = campanas;
        }
        return result;
    }

    // ── Resumen campañas ─────────────────────────────────────
    public Dictionary<string, object> LoadResumenCampanas()
    {
        var filePath = Path.Combine(_dataPath, "resumen_campanas.json");
        if (!File.Exists(filePath)) return new();
        var json = File.ReadAllText(filePath, Encoding.UTF8);
        return JsonSerializer.Deserialize<Dictionary<string, object>>(json) ?? new();
    }

    // ── Primas históricas desde Excel ────────────────────────
    public Dictionary<string, Dictionary<string, double>> LoadPrimasHistoricas()
    {
        var filePath = Path.Combine(_dataPath, "Primas_Totales_SAC_2020-2026.xlsx");
        if (!File.Exists(filePath)) return new();

        var fileBytes = File.ReadAllBytes(filePath);
        using var stream = new MemoryStream(fileBytes);

        IExcelDataReader reader;
        try { reader = ExcelReaderFactory.CreateReader(stream); }
        catch { stream.Position = 0; reader = ExcelReaderFactory.CreateOpenXmlReader(stream); }

        DataSet ds;
        using (reader)
        {
            ds = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = false }
            });
        }

        var sheetMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["SAC 2020-2021"] = "2020-2021", ["SAC 2021-2022"] = "2021-2022",
            ["SAC 2022-2023"] = "2022-2023", ["SAC 2023-2024"] = "2023-2024",
            ["SAC 2024-2025"] = "2024-2025", ["SAC 2025-2026"] = "2025-2026",
        };

        var result = new Dictionary<string, Dictionary<string, double>>();

        foreach (DataTable dt in ds.Tables)
        {
            if (!sheetMap.TryGetValue(dt.TableName.Trim(), out var camp))
                continue;

            var campData = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            for (int i = 1; i < dt.Rows.Count; i++)
            {
                var row = dt.Rows[i];
                var deptoRaw = row[0]?.ToString()?.Trim() ?? "";
                if (string.IsNullOrEmpty(deptoRaw)) continue;
                var depto = NormalizeDepto(deptoRaw);
                if (depto.Contains("TOTAL") || depto.Contains("REGION") || depto.Contains("DEPARTAMENTO"))
                    continue;

                // Prima neta = última columna numérica > 0
                double lastNumeric = 0;
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    var v = TryParseDouble(row[c]);
                    if (v > 0) lastNumeric = v;
                }
                if (lastNumeric > 0)
                    campData[depto] = lastNumeric;
            }
            result[camp] = campData;
        }
        return result;
    }

    // ── Series temporales ────────────────────────────────────
    public SeriesTemporalesDto LoadSeriesTemporales()
    {
        var filePath = Path.Combine(_dataPath, "series_temporales.json");
        if (!File.Exists(filePath)) return new();

        var json = File.ReadAllText(filePath, Encoding.UTF8);
        using var doc = JsonDocument.Parse(json);
        var dto = new SeriesTemporalesDto();

        if (doc.RootElement.TryGetProperty("avisos", out var avisos))
        {
            foreach (var camp in avisos.EnumerateObject())
            {
                var dict = new Dictionary<string, int>();
                foreach (var period in camp.Value.EnumerateObject())
                    dict[period.Name] = period.Value.GetInt32();
                dto.Avisos[camp.Name] = dict;
            }
        }

        if (doc.RootElement.TryGetProperty("indemnizaciones", out var indemn))
        {
            foreach (var camp in indemn.EnumerateObject())
            {
                var dict = new Dictionary<string, IndemnMensualDto>();
                foreach (var period in camp.Value.EnumerateObject())
                {
                    dict[period.Name] = new IndemnMensualDto
                    {
                        N = period.Value.TryGetProperty("n", out var n) ? n.GetInt32() : 0,
                        Monto = period.Value.TryGetProperty("monto", out var m) ? m.GetDouble() : 0,
                    };
                }
                dto.Indemnizaciones[camp.Name] = dict;
            }
        }
        return dto;
    }

    // ── Calendario agrícola ──────────────────────────────────
    public Dictionary<string, List<CalendarioCultivoDto>> LoadCalendarioAgricola()
    {
        var filePath = Path.Combine(_dataPath, "calendario_cultivos_historico.json");
        if (!File.Exists(filePath)) return new();

        var json = File.ReadAllText(filePath, Encoding.UTF8);
        using var doc = JsonDocument.Parse(json);
        var result = new Dictionary<string, List<CalendarioCultivoDto>>(StringComparer.OrdinalIgnoreCase);

        foreach (var deptProp in doc.RootElement.EnumerateObject())
        {
            var cultivos = new List<CalendarioCultivoDto>();
            foreach (var cultProp in deptProp.Value.EnumerateObject())
            {
                var layers = cultProp.Value;
                var avisos = layers.TryGetProperty("avisos", out var av) ? av : default;
                var indemnProp = layers.TryGetProperty("indemnizados", out var ind) ? ind : default;
                var ptotalProp = layers.TryGetProperty("perdida_total", out var ptv) ? ptv : default;

                var entry = new CalendarioCultivoDto
                {
                    Cultivo = cultProp.Name,
                    MesesSiembra = GetIntList(avisos, "meses_siembra"),
                    MesesCosecha = GetIntList(avisos, "meses_cosecha"),
                    MesesRiesgo = GetIntList(avisos, "meses_riesgo"),
                    Riesgos = indemnProp.ValueKind != JsonValueKind.Undefined
                        ? GetStringList(indemnProp, "riesgos")
                        : GetStringList(avisos, "riesgos"),
                    GruposClimaticos = GetStringList(avisos, "grupos_climaticos"),
                    TotalAvisos = avisos.ValueKind != JsonValueKind.Undefined && avisos.TryGetProperty("total", out var ta) ? ta.GetInt32() : 0,
                    TotalIndemnizados = indemnProp.ValueKind != JsonValueKind.Undefined && indemnProp.TryGetProperty("total", out var ti) ? ti.GetInt32() : 0,
                    TotalPerdidaTotal = ptotalProp.ValueKind != JsonValueKind.Undefined && ptotalProp.TryGetProperty("total", out var tp) ? tp.GetInt32() : 0,
                    RiesgosAvisos = GetStringList(avisos, "riesgos"),
                    RiesgosIndemnizados = indemnProp.ValueKind != JsonValueKind.Undefined ? GetStringList(indemnProp, "riesgos") : new(),
                    RiesgosPerdidaTotal = ptotalProp.ValueKind != JsonValueKind.Undefined ? GetStringList(ptotalProp, "riesgos") : new(),
                };

                // Fallback: si no hay siembra ni cosecha, usar meses_riesgo
                if (entry.MesesSiembra.Count == 0 && entry.MesesCosecha.Count == 0)
                    entry.MesesSiembra = new List<int>(entry.MesesRiesgo);

                cultivos.Add(entry);
            }
            result[deptProp.Name] = cultivos;
        }
        return result;
    }

    private static List<int> GetIntList(JsonElement el, string prop)
    {
        if (el.ValueKind == JsonValueKind.Undefined) return new();
        if (!el.TryGetProperty(prop, out var arr) || arr.ValueKind != JsonValueKind.Array) return new();
        return arr.EnumerateArray().Select(x => x.GetInt32()).ToList();
    }

    private static List<string> GetStringList(JsonElement el, string prop)
    {
        if (el.ValueKind == JsonValueKind.Undefined) return new();
        if (!el.TryGetProperty(prop, out var arr) || arr.ValueKind != JsonValueKind.Array) return new();
        return arr.EnumerateArray().Select(x => x.GetString() ?? "").Where(s => s.Length > 0).ToList();
    }

    private static readonly Dictionary<string, string> CapitalToDepto = new(StringComparer.OrdinalIgnoreCase)
    {
        ["CHACHAPOYAS"] = "AMAZONAS",
        ["HUARAZ"] = "ANCASH",
        ["AREQUIPA"] = "AREQUIPA",
        ["CUSCO"] = "CUSCO",
        ["HUANCAVELICA"] = "HUANCAVELICA",
        ["HUÁNUCO"] = "HUANUCO",
        ["HUANUCO"] = "HUANUCO",
        ["HUANCAYO"] = "JUNIN",
        ["TRUJILLO"] = "LA LIBERTAD",
        ["CHICLAYO"] = "LAMBAYEQUE",
        ["LIMA"] = "LIMA",
        ["IQUITOS"] = "LORETO",
        ["MADRE DE DIOS"] = "MADRE DE DIOS",
        ["PUERTO MALDONADO"] = "MADRE DE DIOS",
        ["CERRO DE PASCO"] = "PASCO",
        ["PIURA"] = "PIURA",
        ["PUNO"] = "PUNO",
        ["MOYOBAMBA"] = "SAN MARTIN",
        ["TACNA"] = "TACNA",
        ["PUCALLPA"] = "UCAYALI",
        ["ABANCAY"] = "APURIMAC",
        ["AYACUCHO"] = "AYACUCHO",
        ["CAJAMARCA"] = "CAJAMARCA",
        ["ICA"] = "ICA",
        ["MOQUEGUA"] = "MOQUEGUA",
        ["TUMBES"] = "TUMBES",
    };

    public List<MateriaAsegurada> LoadMateriaAsegurada()
    {
        var filePath = Path.Combine(_dataPath, "Materia_Asegurada_SAC_2025-2026.xlsx");
        var fileBytes = File.ReadAllBytes(filePath);
        using var stream = new MemoryStream(fileBytes);

        IExcelDataReader reader;
        try
        {
            reader = ExcelReaderFactory.CreateReader(stream);
        }
        catch
        {
            stream.Position = 0;
            reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
        }

        DataSet ds;
        using (reader)
        {
            ds = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = false }
            });
        }

        if (ds.Tables.Count == 0) return new List<MateriaAsegurada>();
        var dt = ds.Tables[0];
        if (dt.Rows.Count == 0) return new List<MateriaAsegurada>();

        // Find header row
        int headerRowIdx = 0;
        for (int i = 0; i < Math.Min(dt.Rows.Count, 10); i++)
        {
            var rowVals = new List<string>();
            for (int c = 0; c < dt.Columns.Count; c++)
                rowVals.Add(dt.Rows[i][c]?.ToString()?.Trim() ?? "");

            if (rowVals.Any(v => v.Equals("Capital", StringComparison.OrdinalIgnoreCase) ||
                               v.Equals("Departamento", StringComparison.OrdinalIgnoreCase)))
            {
                headerRowIdx = i;
                break;
            }
        }

        // Map headers
        var colMap = new Dictionary<int, string>();
        for (int c = 0; c < dt.Columns.Count; c++)
        {
            var h = dt.Rows[headerRowIdx][c]?.ToString()?.Trim().ToUpper() ?? "";
            if (string.IsNullOrEmpty(h)) continue;
            if (h.Contains("DEPARTAMENTO")) colMap[c] = "DEPARTAMENTO";
            else if (h.Contains("CAPITAL")) colMap[c] = "CAPITAL";
            else if (h.Contains("EMPRESA")) colMap[c] = "EMPRESA";
            else if (h.Contains("CULTIVOS")) colMap[c] = "CULTIVOS";
            else if (h.Contains("PRIMA TOTAL")) colMap[c] = "PRIMA_TOTAL";
            else if (h.Contains("PRIMA NETA")) colMap[c] = "PRIMA_NETA";
            else if (h.Contains("SUPERFICIE ASEGURADA")) colMap[c] = "SUP_ASEGURADA";
            else if (h.Contains("PRODUCTORES")) colMap[c] = "PRODUCTORES";
            else if (h.Contains("VALORES")) colMap[c] = "VALORES";
            else if (h.Contains("DISPARADOR")) colMap[c] = "DISPARADOR";
            else if (h.Contains("SUMA ASEGURADA")) colMap[c] = "SUMA_ASEGURADA";
        }

        var result = new List<MateriaAsegurada>();
        for (int i = headerRowIdx + 1; i < dt.Rows.Count; i++)
        {
            var ma = new MateriaAsegurada();
            string capital = "";

            foreach (var kvp in colMap)
            {
                var rawValue = dt.Rows[i][kvp.Key];
                if (rawValue == null || rawValue == DBNull.Value) continue;
                var val = rawValue.ToString()?.Trim() ?? "";
                if (string.IsNullOrEmpty(val)) continue;

                switch (kvp.Value)
                {
                    case "DEPARTAMENTO": ma.Departamento = val.ToUpper(); break;
                    case "CAPITAL": capital = val.ToUpper(); ma.Capital = val; break;
                    case "EMPRESA": ma.EmpresaAseguradora = val; break;
                    case "CULTIVOS": ma.CultivosAsegurados = val; break;
                    case "PRIMA_TOTAL": ma.PrimaTotal = TryParseDouble(rawValue); break;
                    case "PRIMA_NETA": ma.PrimaNeta = TryParseDouble(rawValue); break;
                    case "SUP_ASEGURADA": ma.SuperficieAsegurada = TryParseDouble(rawValue); break;
                    case "PRODUCTORES": ma.ProductoresAsegurados = TryParseDouble(rawValue); break;
                    case "VALORES": ma.ValoresAsegurados = TryParseDouble(rawValue); break;
                    case "DISPARADOR": ma.Disparador = val; break;
                    case "SUMA_ASEGURADA": ma.SumaAseguradaHa = TryParseDouble(rawValue); break;
                }
            }

            if (!string.IsNullOrEmpty(capital) && CapitalToDepto.TryGetValue(capital, out var depto))
            {
                ma.Departamento = depto;
            }

            if (!string.IsNullOrEmpty(ma.Departamento) && ma.Departamento != "TOTAL")
            {
                result.Add(ma);
            }
        }

        return result;
    }

    private static double TryParseDouble(object rawValue)
    {
        if (rawValue == null || rawValue == DBNull.Value) return 0;
        if (rawValue is double d) return d;
        if (rawValue is decimal dec) return (double)dec;
        if (rawValue is int intVal) return intVal;
        if (rawValue is long longVal) return longVal;
        if (rawValue is float f) return f;
        var str = rawValue.ToString()?.Trim() ?? "";
        if (string.IsNullOrEmpty(str) || str == "-") return 0;
        if (double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
            return result;
        return 0;
    }
}
