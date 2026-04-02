namespace Infrastructure.StaticData;

using System.Data;
using System.Globalization;
using System.Text;
using Application.Contracts.Repositories;
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
        using var stream = File.OpenRead(filePath);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        var ds = reader.AsDataSet(new ExcelDataSetConfiguration
        {
            ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = false }
        });

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
