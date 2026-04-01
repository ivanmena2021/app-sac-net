namespace Infrastructure.StaticData;

using Application.Contracts.Repositories;
using Domain.Entities;
using ClosedXML.Excel;
using Microsoft.Extensions.Hosting;

public class StaticDataRepository : IStaticDataRepository
{
    private readonly string _dataPath;

    public StaticDataRepository(IHostEnvironment env)
    {
        _dataPath = Path.Combine(env.ContentRootPath, "wwwroot", "data");
    }

    // Capital -> Departamento mapping (same as Python CAPITAL_TO_DEPTO)
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
        using var workbook = new XLWorkbook(filePath);
        var ws = workbook.Worksheet(1);
        var rows = ws.RangeUsed()?.RowsUsed().ToList() ?? new List<IXLRangeRow>();
        if (rows.Count == 0) return new List<MateriaAsegurada>();

        // Find header row (contains "Capital" or "Departamento")
        int headerRowIdx = 0;
        for (int i = 0; i < Math.Min(rows.Count, 10); i++)
        {
            var vals = rows[i].CellsUsed().Select(c => c.GetString().Trim()).ToList();
            if (vals.Any(v => v.Equals("Capital", StringComparison.OrdinalIgnoreCase) ||
                             v.Equals("Departamento", StringComparison.OrdinalIgnoreCase)))
            {
                headerRowIdx = i;
                break;
            }
        }

        // Map headers
        var headerCells = rows[headerRowIdx].CellsUsed().ToList();
        var colMap = new Dictionary<int, string>();
        foreach (var cell in headerCells)
        {
            var h = cell.GetString().Trim().ToUpper();
            if (h.Contains("DEPARTAMENTO")) colMap[cell.Address.ColumnNumber] = "DEPARTAMENTO";
            else if (h.Contains("CAPITAL")) colMap[cell.Address.ColumnNumber] = "CAPITAL";
            else if (h.Contains("EMPRESA")) colMap[cell.Address.ColumnNumber] = "EMPRESA";
            else if (h.Contains("CULTIVOS")) colMap[cell.Address.ColumnNumber] = "CULTIVOS";
            else if (h.Contains("PRIMA TOTAL")) colMap[cell.Address.ColumnNumber] = "PRIMA_TOTAL";
            else if (h.Contains("PRIMA NETA")) colMap[cell.Address.ColumnNumber] = "PRIMA_NETA";
            else if (h.Contains("SUPERFICIE ASEGURADA")) colMap[cell.Address.ColumnNumber] = "SUP_ASEGURADA";
            else if (h.Contains("PRODUCTORES")) colMap[cell.Address.ColumnNumber] = "PRODUCTORES";
            else if (h.Contains("VALORES")) colMap[cell.Address.ColumnNumber] = "VALORES";
            else if (h.Contains("DISPARADOR")) colMap[cell.Address.ColumnNumber] = "DISPARADOR";
            else if (h.Contains("SUMA ASEGURADA")) colMap[cell.Address.ColumnNumber] = "SUMA_ASEGURADA";
        }

        var result = new List<MateriaAsegurada>();
        for (int i = headerRowIdx + 1; i < rows.Count; i++)
        {
            var ma = new MateriaAsegurada();
            string capital = "";

            foreach (var kvp in colMap)
            {
                var cell = ws.Cell(rows[i].RowNumber(), kvp.Key);
                var val = cell.GetString().Trim();
                if (string.IsNullOrEmpty(val)) continue;

                switch (kvp.Value)
                {
                    case "DEPARTAMENTO": ma.Departamento = val.ToUpper(); break;
                    case "CAPITAL": capital = val.ToUpper(); ma.Capital = val; break;
                    case "EMPRESA": ma.EmpresaAseguradora = val; break;
                    case "CULTIVOS": ma.CultivosAsegurados = val; break;
                    case "PRIMA_TOTAL": ma.PrimaTotal = TryParseDouble(val, cell); break;
                    case "PRIMA_NETA": ma.PrimaNeta = TryParseDouble(val, cell); break;
                    case "SUP_ASEGURADA": ma.SuperficieAsegurada = TryParseDouble(val, cell); break;
                    case "PRODUCTORES": ma.ProductoresAsegurados = TryParseDouble(val, cell); break;
                    case "VALORES": ma.ValoresAsegurados = TryParseDouble(val, cell); break;
                    case "DISPARADOR": ma.Disparador = val; break;
                    case "SUMA_ASEGURADA": ma.SumaAseguradaHa = TryParseDouble(val, cell); break;
                }
            }

            // Map capital to department (the DEPARTAMENTO column is often empty)
            if (!string.IsNullOrEmpty(capital) && CapitalToDepto.TryGetValue(capital, out var depto))
            {
                ma.Departamento = depto;
            }

            // Only add if we have a valid department
            if (!string.IsNullOrEmpty(ma.Departamento) && ma.Departamento != "TOTAL")
            {
                result.Add(ma);
            }
        }

        return result;
    }

    private static double TryParseDouble(string value, IXLCell cell)
    {
        if (string.IsNullOrWhiteSpace(value) || value == "-") return 0;
        if (cell.DataType == XLDataType.Number)
        {
            try { return cell.GetDouble(); } catch { }
        }
        if (double.TryParse(value, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var result))
            return result;
        return 0;
    }
}
