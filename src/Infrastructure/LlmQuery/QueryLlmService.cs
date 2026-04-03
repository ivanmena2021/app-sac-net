namespace Infrastructure.LlmQuery;

using System.Data;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using Application.Contracts.Services;
using Application.DTOs.Responses;
using DuckDB.NET.Data;
using Microsoft.Extensions.Configuration;

public class QueryLlmService : IQueryLlmService
{
    private readonly IConfiguration _config;
    private readonly IHttpClientFactory _httpFactory;
    private const string Model = "claude-sonnet-4-20250514";
    private const string ApiUrl = "https://api.anthropic.com/v1/messages";

    public QueryLlmService(IConfiguration config, IHttpClientFactory httpFactory)
    {
        _config = config;
        _httpFactory = httpFactory;
    }

    public bool IsAvailable()
    {
        var key = _config["Anthropic:ApiKey"]
            ?? Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY");
        return !string.IsNullOrEmpty(key);
    }

    private string GetApiKey() =>
        _config["Anthropic:ApiKey"]
        ?? Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY")
        ?? throw new InvalidOperationException("ANTHROPIC_API_KEY no configurada.");

    public async Task<QueryLlmResult> ProcessQueryAsync(string question, DatosNacionalesDto datos)
    {
        string apiKey;
        try { apiKey = GetApiKey(); }
        catch (Exception ex) { return new QueryLlmResult { Error = ex.Message }; }

        DuckDBConnection conn;
        string schema;
        try { (conn, schema) = LoadToDuckDb(datos); }
        catch (Exception ex) { return new QueryLlmResult { Error = $"Error DuckDB: {ex.Message}" }; }

        try
        {
            // Generate SQL
            string sql;
            try { sql = await CallClaude(apiKey, SystemSql, $"Esquema:\n{schema}\n\nPregunta: {question}", 1024); }
            catch (Exception ex) { return new QueryLlmResult { Error = $"Error SQL: {ex.Message}" }; }
            sql = CleanSql(sql);

            // Execute
            var (result, err) = ExecuteSql(conn, sql);
            if (err != null)
            {
                // Retry
                try
                {
                    sql = await CallClaude(apiKey, SystemSql,
                        $"Esquema:\n{schema}\n\nPregunta: {question}\n\nSQL fallida: {sql}\nError: {err}\nGenera SQL corregida.", 1024);
                    sql = CleanSql(sql);
                    (result, err) = ExecuteSql(conn, sql);
                }
                catch { }
            }
            if (err != null) return new QueryLlmResult { Sql = sql, Error = $"Error SQL: {err}" };

            // Summary
            var (summary, summaryText) = ComputeVerifiedSummary(conn, sql, result);

            // Prose
            string? prose = null;
            try
            {
                var resultText = TableToText(result, 50);
                var userMsg = $"Fecha de corte: {datos.FechaCorte}\nPregunta: {question}\n\n" +
                    $"TABLA DETALLADA ({result?.Rows.Count ?? 0} filas):\n{resultText}\n\n{summaryText}\n\n" +
                    "Usa cifras del RESUMEN VERIFICADO para totales. Redacta texto profesional.";
                prose = await CallClaude(apiKey, SystemProse, userMsg, 2048);
            }
            catch (Exception ex)
            {
                return new QueryLlmResult { Sql = sql, Data = TableToList(result), Summary = summary, Error = $"Error prosa: {ex.Message}" };
            }

            return new QueryLlmResult { Prose = prose, Sql = sql, Data = TableToList(result), Summary = summary };
        }
        finally { conn.Dispose(); }
    }

    // ── Claude API call via HttpClient ───────────────────────────────

    private async Task<string> CallClaude(string apiKey, string system, string userMessage, int maxTokens)
    {
        using var http = _httpFactory.CreateClient();
        http.Timeout = TimeSpan.FromSeconds(60);

        var request = new HttpRequestMessage(HttpMethod.Post, ApiUrl);
        request.Headers.Add("x-api-key", apiKey);
        request.Headers.Add("anthropic-version", "2023-06-01");

        var body = new
        {
            model = Model,
            max_tokens = maxTokens,
            system = system,
            messages = new[] { new { role = "user", content = userMessage } },
        };
        request.Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");

        var response = await http.SendAsync(request);
        var json = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
            throw new Exception($"Anthropic API {response.StatusCode}: {json[..Math.Min(300, json.Length)]}");

        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.GetProperty("content")[0].GetProperty("text").GetString() ?? "";
    }

    private static string CleanSql(string sql)
    {
        sql = sql.Trim();
        sql = Regex.Replace(sql, @"^```(?:sql)?\s*", "");
        sql = Regex.Replace(sql, @"\s*```$", "");
        return sql.Trim();
    }

    // ── DuckDB ──────────────────────────────────────────────────────

    private static (DuckDBConnection conn, string schema) LoadToDuckDb(DatosNacionalesDto datos)
    {
        var conn = new DuckDBConnection("DataSource=:memory:");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"CREATE TABLE avisos (
            DEPARTAMENTO VARCHAR, PROVINCIA VARCHAR, DISTRITO VARCHAR, SECTOR_ESTADISTICO VARCHAR,
            TIPO_CULTIVO VARCHAR, TIPO_SINIESTRO VARCHAR, ESTADO_INSPECCION VARCHAR, DICTAMEN VARCHAR,
            EMPRESA VARCHAR, CODIGO_AVISO VARCHAR,
            INDEMNIZACION DOUBLE, MONTO_DESEMBOLSADO DOUBLE, SUP_INDEMNIZADA DOUBLE,
            N_PRODUCTORES DOUBLE, SUP_AFECTADA DOUBLE, SUP_PERDIDA DOUBLE,
            FECHA_AVISO DATE, FECHA_SINIESTRO DATE, FECHA_ATENCION DATE, FECHA_DESEMBOLSO DATE
        )";
        cmd.ExecuteNonQuery();

        var deptoEmpresa = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var m in datos.Materia)
        {
            if (string.IsNullOrEmpty(m.Departamento) || string.IsNullOrEmpty(m.EmpresaAseguradora)) continue;
            var emp = m.EmpresaAseguradora.ToUpper();
            if (emp.Contains("POSITIVA")) emp = "LA POSITIVA";
            else if (emp.Contains("RIMAC") || emp.Contains("RÍMAC")) emp = "RIMAC";
            deptoEmpresa[m.Departamento.ToUpper()] = emp;
        }

        foreach (var s in datos.Midagri)
        {
            var empresa = deptoEmpresa.GetValueOrDefault(s.Departamento?.ToUpper() ?? "", "OTROS");
            using var ins = conn.CreateCommand();
            ins.CommandText = "INSERT INTO avisos VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20)";
            var ps = new object?[] {
                s.Departamento, s.Provincia, s.Distrito, s.SectorEstadistico,
                s.TipoCultivo, s.TipoSiniestro, s.EstadoInspeccion, s.Dictamen,
                empresa, s.CodigoAviso,
                s.Indemnizacion, s.MontoDesembolsado, s.SupIndemnizada,
                s.NProductores, s.SupAfectada, s.SupPerdida,
                s.FechaAviso.HasValue ? (object)s.FechaAviso.Value : DBNull.Value,
                s.FechaSiniestro.HasValue ? (object)s.FechaSiniestro.Value : DBNull.Value,
                s.FechaAtencion.HasValue ? (object)s.FechaAtencion.Value : DBNull.Value,
                s.FechaDesembolso.HasValue ? (object)s.FechaDesembolso.Value : DBNull.Value,
            };
            for (int i = 0; i < ps.Length; i++)
            {
                var p = ins.CreateParameter();
                p.ParameterName = $"${i + 1}";
                p.Value = ps[i] ?? DBNull.Value;
                ins.Parameters.Add(p);
            }
            ins.ExecuteNonQuery();
        }

        var schema = GenerateSchema(conn);
        return (conn, schema);
    }

    private static string GenerateSchema(DuckDBConnection conn)
    {
        var sb = new StringBuilder();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT COUNT(*) FROM avisos";
        var count = Convert.ToInt64(cmd.ExecuteScalar());
        sb.AppendLine($"Tabla: avisos ({count:N0} filas)");
        sb.AppendLine("Columnas:");
        cmd.CommandText = "DESCRIBE avisos";
        using (var r = cmd.ExecuteReader())
            while (r.Read()) sb.AppendLine($"    {r.GetString(0)} ({r.GetString(1)})");

        try { cmd.CommandText = "SELECT DISTINCT DEPARTAMENTO FROM avisos ORDER BY 1"; using var r2 = cmd.ExecuteReader(); var v = new List<string>(); while (r2.Read()) if (!r2.IsDBNull(0)) v.Add(r2.GetString(0)); sb.AppendLine($"  Departamentos: {string.Join(", ", v)}"); } catch { }
        try { cmd.CommandText = "SELECT DISTINCT TIPO_SINIESTRO FROM avisos WHERE TIPO_SINIESTRO IS NOT NULL ORDER BY 1"; using var r3 = cmd.ExecuteReader(); var v = new List<string>(); while (r3.Read()) v.Add(r3.GetString(0)); sb.AppendLine($"  Tipos: {string.Join(", ", v)}"); } catch { }
        try { cmd.CommandText = "SELECT COUNT(DISTINCT DEPARTAMENTO),COUNT(DISTINCT PROVINCIA),COUNT(DISTINCT DISTRITO) FROM avisos"; using var r4 = cmd.ExecuteReader(); if (r4.Read()) sb.AppendLine($"  Niveles: {r4.GetInt64(0)} deptos, {r4.GetInt64(1)} provs, {r4.GetInt64(2)} dists"); } catch { }

        return sb.ToString();
    }

    private static (DataTable? result, string? error) ExecuteSql(DuckDBConnection conn, string sql)
    {
        try
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            using var reader = cmd.ExecuteReader();
            var dt = new DataTable();
            dt.Load(reader);
            return (dt, null);
        }
        catch (Exception ex) { return (null, ex.Message); }
    }

    // ── Verified Summary ────────────────────────────────────────────

    private static (Dictionary<string, object>? summary, string text) ComputeVerifiedSummary(DuckDBConnection conn, string sql, DataTable? resultTable)
    {
        if (resultTable == null || resultTable.Rows.Count == 0) return (null, "");
        var where = ExtractWhereClause(sql);
        var summarySql = $@"SELECT COUNT(*),ROUND(SUM(COALESCE(INDEMNIZACION,0)),2),ROUND(SUM(COALESCE(MONTO_DESEMBOLSADO,0)),2),
            ROUND(SUM(COALESCE(SUP_INDEMNIZADA,0)),2),COALESCE(SUM(COALESCE(N_PRODUCTORES,0)),0),
            SUM(CASE WHEN UPPER(COALESCE(ESTADO_INSPECCION,''))='CERRADO' THEN 1 ELSE 0 END),
            ROUND(SUM(CASE WHEN UPPER(COALESCE(ESTADO_INSPECCION,''))='CERRADO' THEN 1 ELSE 0 END)*100.0/NULLIF(COUNT(*),0),1),
            ROUND(SUM(COALESCE(MONTO_DESEMBOLSADO,0))*100.0/NULLIF(SUM(COALESCE(INDEMNIZACION,0)),0),1)
            FROM avisos {where}";
        try
        {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = summarySql;
            using var r = cmd.ExecuteReader();
            if (!r.Read()) return (null, "");
            var s = new Dictionary<string, object>
            {
                ["total_avisos"] = r.IsDBNull(0) ? 0L : Convert.ToInt64(r.GetValue(0)),
                ["total_indemnizacion"] = r.IsDBNull(1) ? 0.0 : Convert.ToDouble(r.GetValue(1)),
                ["total_desembolso"] = r.IsDBNull(2) ? 0.0 : Convert.ToDouble(r.GetValue(2)),
                ["total_ha_indemnizadas"] = r.IsDBNull(3) ? 0.0 : Convert.ToDouble(r.GetValue(3)),
                ["total_productores"] = r.IsDBNull(4) ? 0L : Convert.ToInt64(r.GetValue(4)),
                ["pct_evaluacion"] = r.IsDBNull(6) ? 0.0 : Convert.ToDouble(r.GetValue(6)),
                ["pct_desembolso"] = r.IsDBNull(7) ? 0.0 : Convert.ToDouble(r.GetValue(7)),
            };
            var t = $"RESUMEN VERIFICADO:\nAvisos: {s["total_avisos"]:N0}\nIndemnizacion: S/ {s["total_indemnizacion"]:N2}\n" +
                $"Desembolso: S/ {s["total_desembolso"]:N2}\nHa: {s["total_ha_indemnizadas"]:N2}\n" +
                $"Productores: {s["total_productores"]:N0}\n% Evaluacion: {s["pct_evaluacion"]}%\n% Desembolso: {s["pct_desembolso"]}%";
            return (s, t);
        }
        catch { return (null, ""); }
    }

    private static string ExtractWhereClause(string sql)
    {
        var u = sql.ToUpper();
        int d = 0, ws = -1;
        for (int i = 0; i < u.Length; i++)
        {
            if (u[i] == '(') d++; else if (u[i] == ')') d--;
            else if (d == 0 && i + 5 <= u.Length && u.Substring(i, 5) == "WHERE") { ws = i; break; }
        }
        if (ws == -1) return "";
        int we = sql.Length;
        foreach (var kw in new[] { "GROUP BY", "ORDER BY", "LIMIT", "HAVING" })
        {
            var p = u.IndexOf(kw, ws + 5, StringComparison.Ordinal);
            if (p != -1 && p < we) we = p;
        }
        return sql[ws..we].Trim();
    }

    // ── Prompts ─────────────────────────────────────────────────────

    private static readonly string SystemSql = @"Eres experto SQL del SAC Peru. Convierte preguntas a SQL para DuckDB.
Tabla ""avisos"": DEPARTAMENTO,PROVINCIA,DISTRITO,SECTOR_ESTADISTICO,TIPO_CULTIVO,TIPO_SINIESTRO,ESTADO_INSPECCION,DICTAMEN,EMPRESA,CODIGO_AVISO,INDEMNIZACION,MONTO_DESEMBOLSADO,SUP_INDEMNIZADA,N_PRODUCTORES,FECHA_AVISO,FECHA_SINIESTRO,FECHA_ATENCION,FECHA_DESEMBOLSO.
NUNCA uses SUP_AFECTADA. EMPRESA: 'LA POSITIVA' o 'RIMAC'. Fechas tipo DATE, usa YEAR()/MONTH().
""lluvias""->IN('INUNDACION','LLUVIAS EXCESIVAS','HUAYCO','DESLIZAMIENTO'). ""frio""->IN('HELADA','FRIAJE','NIEVE').
Genera UNA SQL de detalle con ORDER BY. Devuelve SOLO SQL sin backticks.";

    private static readonly string SystemProse = @"Eres redactor profesional de MIDAGRI Peru. Transforma datos SAC 2025-2026 en texto profesional.
Espanol formal, parrafos fluidos, sin markdown. Montos: S/ 1,234.89. Ha para hectareas.
NUNCA reportes SUP_AFECTADA. Usa cifras del RESUMEN VERIFICADO para totales.
Termina con: Fuente: Direccion de Seguro y Fomento del Financiamiento Agrario - MIDAGRI, SAC 2025-2026.";

    // ── Helpers ──────────────────────────────────────────────────────

    private static string TableToText(DataTable? dt, int maxRows)
    {
        if (dt == null || dt.Rows.Count == 0) return "(Sin resultados)";
        var sb = new StringBuilder();
        var cols = new List<string>();
        foreach (DataColumn c in dt.Columns) cols.Add(c.ColumnName);
        sb.AppendLine(string.Join("\t", cols));
        for (int i = 0; i < Math.Min(maxRows, dt.Rows.Count); i++)
        {
            var vals = new List<string>();
            foreach (DataColumn c in dt.Columns) vals.Add(dt.Rows[i][c]?.ToString() ?? "");
            sb.AppendLine(string.Join("\t", vals));
        }
        return sb.ToString();
    }

    private static List<Dictionary<string, object>>? TableToList(DataTable? dt)
    {
        if (dt == null) return null;
        var list = new List<Dictionary<string, object>>();
        foreach (DataRow row in dt.Rows)
        {
            var dict = new Dictionary<string, object>();
            foreach (DataColumn col in dt.Columns)
                dict[col.ColumnName] = row[col] == DBNull.Value ? "" : row[col];
            list.Add(dict);
        }
        return list;
    }
}
