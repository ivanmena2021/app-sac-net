namespace Infrastructure.ExcelReader;

using System.Globalization;
using Application.Contracts.Repositories;
using Domain.Entities;
using ClosedXML.Excel;

public class ExcelReaderRepository : IExcelReaderRepository
{
    public List<Siniestro> ReadMidagriExcel(Stream fileStream)
    {
        using var workbook = new XLWorkbook(fileStream);
        var ws = workbook.Worksheet(1);
        var rows = ws.RangeUsed()?.RowsUsed().ToList() ?? new List<IXLRangeRow>();
        if (rows.Count == 0) return new List<Siniestro>();

        // Find header row (contains "CAMPAÑA" or "CÓDIGO DE AVISO")
        int headerRowIdx = 0;
        for (int i = 0; i < Math.Min(rows.Count, 5); i++)
        {
            var vals = rows[i].CellsUsed().Select(c => c.GetString().Trim().ToUpper()).ToList();
            if (vals.Any(v => v.Contains("CAMPAÑA") || v.Contains("CODIGO DE AVISO") || v.Contains("CÓDIGO DE AVISO")))
            {
                headerRowIdx = i;
                break;
            }
        }

        // Map header columns
        var headerCells = rows[headerRowIdx].CellsUsed().ToList();
        var colMap = new Dictionary<int, string>();
        foreach (var cell in headerCells)
        {
            var colName = MapColumnName(cell.GetString().Trim().ToUpper());
            if (!string.IsNullOrEmpty(colName))
                colMap[cell.Address.ColumnNumber] = colName;
        }

        // Read data rows
        var result = new List<Siniestro>();
        for (int i = headerRowIdx + 1; i < rows.Count; i++)
        {
            var siniestro = new Siniestro();
            bool hasData = false;

            foreach (var kvp in colMap)
            {
                var cell = ws.Cell(rows[i].RowNumber(), kvp.Key);
                var value = cell.GetString().Trim();
                if (string.IsNullOrEmpty(value) || value == "-" || value.ToUpper() == "NAN") continue;
                hasData = true;
                SetSiniestroProperty(siniestro, kvp.Value, value, cell);
            }

            if (hasData && !string.IsNullOrEmpty(siniestro.Departamento) &&
                siniestro.Departamento.ToUpper() != "NAN" && siniestro.Departamento.ToUpper() != "NONE")
            {
                siniestro.Departamento = siniestro.Departamento.Trim().ToUpper();
                siniestro.TipoSiniestro = siniestro.TipoSiniestro.Trim().ToUpper();
                result.Add(siniestro);
            }
        }

        return result;
    }

    public List<Siniestro> ReadSiniestrosExcel(Stream fileStream)
    {
        // Same logic as ReadMidagriExcel - the normalization is identical
        return ReadMidagriExcel(fileStream);
    }

    private static string MapColumnName(string header)
    {
        // Port ALL the column mappings from Python _normalize_midagri
        if (header.Contains("CAMPAÑA")) return "CAMPANA";
        if (header.Contains("CÓDIGO DE AVISO") || header.Contains("CODIGO DE AVISO")) return "CODIGO_AVISO";
        if (header == "DEPARTAMENTO") return "DEPARTAMENTO";
        if (header == "PROVINCIA") return "PROVINCIA";
        if (header == "DISTRITO") return "DISTRITO";
        if (header.Contains("SECTOR")) return "SECTOR_ESTADISTICO";
        if (header.Contains("TIPO DE CULTIVO") || header.Contains("TIPO CULTIVO")) return "TIPO_CULTIVO";
        if (header.Contains("FENOLOG")) return "FENOLOGIA";
        if (header.Contains("FECHA DE SIEMBRA") || header.Contains("FECHA SIEMBRA")) return "FECHA_SIEMBRA";
        if (header.Contains("FECHA DE COSECHA") || header.Contains("FECHA COSECHA")) return "FECHA_COSECHA";
        if (header.Contains("SUPERFICIE SEMBRADA")) return "SUP_SEMBRADA";
        if (header.Contains("SUPERFICIE ASEGURADA")) return "SUP_ASEGURADA";
        if (header.Contains("TIPO DE SINIESTRO") || header.Contains("TIPO SINIESTRO")) return "TIPO_SINIESTRO";
        if (header.Contains("FECHA DE SINIESTRO") || header.Contains("FECHA SINIESTRO")) return "FECHA_SINIESTRO";
        if (header.Contains("FECHA DE AVISO") || header.Contains("FECHA AVISO")) return "FECHA_AVISO";
        if (header.Contains("FECHA DE ATENCIÓN") || header.Contains("FECHA ATENCION") || header.Contains("FECHA DE ATENCION")) return "FECHA_ATENCION";
        if (header.Contains("ESTADO SINIESTRO")) return "ESTADO_SINIESTRO";
        if (header.Contains("ESTADO INSPECCION") || header.Contains("ESTADO INSPECCIÓN")) return "ESTADO_INSPECCION";
        if (header.Contains("PRIMA NETA")) return "PRIMA_NETA_DPTO";
        if (header.Contains("TIPO DE COBERTURA") || header.Contains("TIPO COBERTURA")) return "TIPO_COBERTURA";
        if (header.Contains("SUPERFICIE AFECTADA")) return "SUP_AFECTADA";
        if (header.Contains("SUPERFICIE PERDIDA")) return "SUP_PERDIDA";
        if (header.Contains("DICTAMEN")) return "DICTAMEN";
        if (header.Contains("SUPERFICIE INDEMNIZADA")) return "SUP_INDEMNIZADA";
        if (header == "INDEMNIZACIÓN" || header.Contains("INDEMNIZACI")) return "INDEMNIZACION";
        if (header.Contains("MONTO DESEMBOLSADO")) return "MONTO_DESEMBOLSADO";
        if (header.Contains("SUPERFICIE DESEMBOLSO")) return "SUP_DESEMBOLSO";
        if (header.Contains("PRODUCTORES") || header.Contains("N° DE PRODUCTORES")) return "N_PRODUCTORES";
        if (header.Contains("CÓDIGO DE PADRÓN") || header.Contains("CODIGO DE PADRON")) return "CODIGO_PADRON";
        if (header.Contains("FECHA DE ENVIO") || header.Contains("FECHA ENVIO")) return "FECHA_ENVIO_DRAS";
        if (header.Contains("FECHA VALIDACI")) return "FECHA_VALIDACION";
        if (header.Contains("FECHA DESEMBOLSO")) return "FECHA_DESEMBOLSO";
        if (header.Contains("PRIORIZADO")) return "PRIORIZADO";
        if (header.Contains("OBSERVACI")) return "OBSERVACION";
        return string.Empty;
    }

    private static void SetSiniestroProperty(Siniestro s, string colName, string value, IXLCell cell)
    {
        switch (colName)
        {
            case "CAMPANA": s.Campana = value; break;
            case "CODIGO_AVISO": s.CodigoAviso = value; break;
            case "DEPARTAMENTO": s.Departamento = value; break;
            case "PROVINCIA": s.Provincia = value; break;
            case "DISTRITO": s.Distrito = value; break;
            case "SECTOR_ESTADISTICO": s.SectorEstadistico = value; break;
            case "TIPO_CULTIVO": s.TipoCultivo = value; break;
            case "FENOLOGIA": s.Fenologia = value; break;
            case "FECHA_SIEMBRA": s.FechaSiembra = TryParseDate(value, cell); break;
            case "FECHA_COSECHA": s.FechaCosecha = TryParseDate(value, cell); break;
            case "SUP_SEMBRADA": s.SupSembrada = TryParseDouble(value); break;
            case "SUP_ASEGURADA": s.SupAsegurada = TryParseDouble(value); break;
            case "TIPO_SINIESTRO": s.TipoSiniestro = value; break;
            case "FECHA_SINIESTRO": s.FechaSiniestro = TryParseDate(value, cell); break;
            case "FECHA_AVISO": s.FechaAviso = TryParseDate(value, cell); break;
            case "FECHA_ATENCION": s.FechaAtencion = TryParseDate(value, cell); break;
            case "ESTADO_SINIESTRO": s.EstadoSiniestro = value; break;
            case "ESTADO_INSPECCION": s.EstadoInspeccion = value; break;
            case "PRIMA_NETA_DPTO": s.PrimaNetaDpto = TryParseDouble(value); break;
            case "TIPO_COBERTURA": s.TipoCobertura = value; break;
            case "SUP_AFECTADA": s.SupAfectada = TryParseDouble(value); break;
            case "SUP_PERDIDA": s.SupPerdida = TryParseDouble(value); break;
            case "DICTAMEN": s.Dictamen = value; break;
            case "SUP_INDEMNIZADA": s.SupIndemnizada = TryParseDouble(value); break;
            case "INDEMNIZACION": s.Indemnizacion = TryParseDouble(value); break;
            case "MONTO_DESEMBOLSADO": s.MontoDesembolsado = TryParseDouble(value); break;
            case "SUP_DESEMBOLSO": s.SupDesembolso = TryParseDouble(value); break;
            case "N_PRODUCTORES": s.NProductores = TryParseDouble(value); break;
            case "CODIGO_PADRON": s.CodigoPadron = value; break;
            case "FECHA_ENVIO_DRAS": s.FechaEnvioDras = TryParseDate(value, cell); break;
            case "FECHA_VALIDACION": s.FechaValidacion = TryParseDate(value, cell); break;
            case "FECHA_DESEMBOLSO": s.FechaDesembolso = TryParseDate(value, cell); break;
            case "PRIORIZADO": s.Priorizado = value; break;
            case "OBSERVACION": s.Observacion = value; break;
        }
    }

    private static double TryParseDouble(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value == "-") return 0;
        if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
            return result;
        return 0;
    }

    private static DateTime? TryParseDate(string value, IXLCell cell)
    {
        // Try cell's date value first
        if (cell.DataType == XLDataType.DateTime)
        {
            try { return cell.GetDateTime(); } catch { }
        }
        // Try parsing string
        if (DateTime.TryParseExact(value, new[] { "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd", "MM/dd/yyyy" },
            CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
            return dt;
        if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
            return dt;
        return null;
    }
}
