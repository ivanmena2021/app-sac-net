namespace Infrastructure.ExcelReader;

using System.Data;
using System.Globalization;
using System.Text;
using Application.Contracts.Repositories;
using Domain.Entities;
using ExcelDataReader;

public class ExcelReaderRepository : IExcelReaderRepository
{
    static ExcelReaderRepository()
    {
        // Required for ExcelDataReader on .NET Core
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public List<Siniestro> ReadMidagriExcel(Stream fileStream)
    {
        return ReadExcel(fileStream);
    }

    public List<Siniestro> ReadSiniestrosExcel(Stream fileStream)
    {
        return ReadExcel(fileStream);
    }

    private static List<Siniestro> ReadExcel(Stream fileStream)
    {
        // Copy to MemoryStream if not seekable (Blazor uploads)
        Stream workStream = fileStream;
        if (!fileStream.CanSeek)
        {
            var ms = new MemoryStream();
            fileStream.CopyTo(ms);
            ms.Position = 0;
            workStream = ms;
        }
        else
        {
            fileStream.Position = 0;
        }

        using var reader = ExcelReaderFactory.CreateReader(workStream);
        var ds = reader.AsDataSet(new ExcelDataSetConfiguration
        {
            ConfigureDataTable = _ => new ExcelDataTableConfiguration
            {
                UseHeaderRow = false // We'll detect header ourselves
            }
        });

        if (ds.Tables.Count == 0) return new List<Siniestro>();
        var dt = ds.Tables[0];
        if (dt.Rows.Count == 0) return new List<Siniestro>();

        // Find header row (contains "CAMPAÑA" or "CÓDIGO DE AVISO" or "DEPARTAMENTO")
        int headerRowIdx = 0;
        for (int i = 0; i < Math.Min(dt.Rows.Count, 10); i++)
        {
            var rowVals = new List<string>();
            for (int c = 0; c < dt.Columns.Count; c++)
            {
                var v = dt.Rows[i][c]?.ToString()?.Trim().ToUpper() ?? "";
                rowVals.Add(v);
            }
            if (rowVals.Any(v => v.Contains("CAMPAÑA") || v.Contains("CODIGO DE AVISO") ||
                                v.Contains("CÓDIGO DE AVISO") || v == "DEPARTAMENTO"))
            {
                headerRowIdx = i;
                break;
            }
        }

        // Map header columns
        var colMap = new Dictionary<int, string>();
        for (int c = 0; c < dt.Columns.Count; c++)
        {
            var headerVal = dt.Rows[headerRowIdx][c]?.ToString()?.Trim().ToUpper() ?? "";
            if (string.IsNullOrEmpty(headerVal)) continue;
            var mapped = MapColumnName(headerVal);
            if (!string.IsNullOrEmpty(mapped) && !colMap.ContainsValue(mapped))
                colMap[c] = mapped;
        }

        // Read data rows
        var result = new List<Siniestro>();
        for (int i = headerRowIdx + 1; i < dt.Rows.Count; i++)
        {
            var siniestro = new Siniestro();
            bool hasData = false;

            foreach (var kvp in colMap)
            {
                var cellValue = dt.Rows[i][kvp.Key];
                if (cellValue == null || cellValue == DBNull.Value) continue;

                var strValue = cellValue.ToString()?.Trim() ?? "";
                if (string.IsNullOrEmpty(strValue) || strValue == "-" ||
                    strValue.Equals("NAN", StringComparison.OrdinalIgnoreCase)) continue;

                hasData = true;
                SetSiniestroProperty(siniestro, kvp.Value, strValue, cellValue);
            }

            if (hasData && !string.IsNullOrEmpty(siniestro.Departamento) &&
                !siniestro.Departamento.Equals("NAN", StringComparison.OrdinalIgnoreCase) &&
                !siniestro.Departamento.Equals("NONE", StringComparison.OrdinalIgnoreCase))
            {
                siniestro.Departamento = siniestro.Departamento.Trim().ToUpper();
                siniestro.TipoSiniestro = (siniestro.TipoSiniestro ?? "").Trim().ToUpper();
                result.Add(siniestro);
            }
        }

        return result;
    }

    private static string MapColumnName(string header)
    {
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
        if (header.Contains("FECHA PROGRAMACION") || header.Contains("FECHA DE PROGRAMACION")) return "FECHA_PROGRAMACION";
        if (header.Contains("FECHA AJUSTE") || header.Contains("FECHA DE AJUSTE")) return "FECHA_AJUSTE";
        if (header.Contains("REPROGRAMACION") || header.Contains("REPROGRAMACIÓN")) return "FECHA_REPROGRAMACION";
        return string.Empty;
    }

    private static void SetSiniestroProperty(Siniestro s, string colName, string strValue, object rawValue)
    {
        switch (colName)
        {
            case "CAMPANA": s.Campana = strValue; break;
            case "CODIGO_AVISO": s.CodigoAviso = strValue; break;
            case "DEPARTAMENTO": s.Departamento = strValue; break;
            case "PROVINCIA": s.Provincia = strValue; break;
            case "DISTRITO": s.Distrito = strValue; break;
            case "SECTOR_ESTADISTICO": s.SectorEstadistico = strValue; break;
            case "TIPO_CULTIVO": s.TipoCultivo = strValue; break;
            case "FENOLOGIA": s.Fenologia = strValue; break;
            case "FECHA_SIEMBRA": s.FechaSiembra = TryParseDate(strValue, rawValue); break;
            case "FECHA_COSECHA": s.FechaCosecha = TryParseDate(strValue, rawValue); break;
            case "SUP_SEMBRADA": s.SupSembrada = TryParseDouble(strValue, rawValue); break;
            case "SUP_ASEGURADA": s.SupAsegurada = TryParseDouble(strValue, rawValue); break;
            case "TIPO_SINIESTRO": s.TipoSiniestro = strValue; break;
            case "FECHA_SINIESTRO": s.FechaSiniestro = TryParseDate(strValue, rawValue); break;
            case "FECHA_AVISO": s.FechaAviso = TryParseDate(strValue, rawValue); break;
            case "FECHA_ATENCION": s.FechaAtencion = TryParseDate(strValue, rawValue); break;
            case "ESTADO_SINIESTRO": s.EstadoSiniestro = strValue; break;
            case "ESTADO_INSPECCION": s.EstadoInspeccion = strValue; break;
            case "PRIMA_NETA_DPTO": s.PrimaNetaDpto = TryParseDouble(strValue, rawValue); break;
            case "TIPO_COBERTURA": s.TipoCobertura = strValue; break;
            case "SUP_AFECTADA": s.SupAfectada = TryParseDouble(strValue, rawValue); break;
            case "SUP_PERDIDA": s.SupPerdida = TryParseDouble(strValue, rawValue); break;
            case "DICTAMEN": s.Dictamen = strValue; break;
            case "SUP_INDEMNIZADA": s.SupIndemnizada = TryParseDouble(strValue, rawValue); break;
            case "INDEMNIZACION": s.Indemnizacion = TryParseDouble(strValue, rawValue); break;
            case "MONTO_DESEMBOLSADO": s.MontoDesembolsado = TryParseDouble(strValue, rawValue); break;
            case "SUP_DESEMBOLSO": s.SupDesembolso = TryParseDouble(strValue, rawValue); break;
            case "N_PRODUCTORES": s.NProductores = TryParseDouble(strValue, rawValue); break;
            case "CODIGO_PADRON": s.CodigoPadron = strValue; break;
            case "FECHA_ENVIO_DRAS": s.FechaEnvioDras = TryParseDate(strValue, rawValue); break;
            case "FECHA_VALIDACION": s.FechaValidacion = TryParseDate(strValue, rawValue); break;
            case "FECHA_DESEMBOLSO": s.FechaDesembolso = TryParseDate(strValue, rawValue); break;
            case "PRIORIZADO": s.Priorizado = strValue; break;
            case "OBSERVACION": s.Observacion = strValue; break;
        }
    }

    private static double TryParseDouble(string strValue, object rawValue)
    {
        if (string.IsNullOrWhiteSpace(strValue) || strValue == "-") return 0;
        // Try raw value first (preserves numeric precision)
        if (rawValue is double d) return d;
        if (rawValue is decimal dec) return (double)dec;
        if (rawValue is int intVal) return intVal;
        if (rawValue is long longVal) return longVal;
        if (rawValue is float f) return f;
        // Parse string
        if (double.TryParse(strValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
            return result;
        // Try with comma as decimal separator
        if (double.TryParse(strValue.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out result))
            return result;
        return 0;
    }

    private static DateTime? TryParseDate(string strValue, object rawValue)
    {
        // Try raw value first (ExcelDataReader returns DateTime directly for date cells)
        if (rawValue is DateTime dt) return dt;
        if (string.IsNullOrWhiteSpace(strValue)) return null;
        // Try common formats
        if (DateTime.TryParseExact(strValue,
            new[] { "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd", "MM/dd/yyyy", "dd-MM-yyyy",
                    "dd/MM/yyyy HH:mm:ss", "d/M/yyyy H:mm:ss", "yyyy-MM-dd HH:mm:ss" },
            CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed))
            return parsed;
        if (DateTime.TryParse(strValue, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed))
            return parsed;
        return null;
    }
}
