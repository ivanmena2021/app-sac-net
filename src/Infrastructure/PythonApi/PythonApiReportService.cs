namespace Infrastructure.PythonApi;

using System.Net.Http;
using System.Net.Http.Headers;
using Application.Contracts.Services;
using Application.DTOs.Requests;
using Application.DTOs.Responses;
using Microsoft.Extensions.Configuration;

/// <summary>
/// Calls the Python FastAPI microservice for document generation.
/// Replaces the local .NET generators with the professional Python originals.
/// Uses IUploadedFileStore (scoped per-circuit) for multi-user safety.
/// </summary>
public class PythonApiReportService : IWordReportService, IExcelReportService,
    IPdfReportService, IExcelEnhancedService, IWordOperatividadService, IPptReportService
{
    private readonly HttpClient _httpClient;
    private readonly IUploadedFileStore _fileStore;
    private readonly string _baseUrl;

    public PythonApiReportService(HttpClient httpClient, IConfiguration configuration,
        IUploadedFileStore fileStore)
    {
        _httpClient = httpClient;
        _fileStore = fileStore;
        _baseUrl = configuration["PythonApi:BaseUrl"] ?? "http://localhost:8000";
        _httpClient.Timeout = TimeSpan.FromMinutes(5); // PPT generation can be slow
    }

    // ─── IWordReportService ───
    public byte[] GenerateNacionalDocx(DatosNacionalesDto datos)
        => CallPythonApi("word-nacional").GetAwaiter().GetResult();

    public byte[] GenerateDepartamentalDocx(DatosDepartamentalDto datos)
        => CallPythonApi("word-departamental", datos.Departamento).GetAwaiter().GetResult();

    // ─── IExcelReportService ───
    public byte[] GenerateReporteEme(DatosNacionalesDto datos)
        => CallPythonApi("excel-eme").GetAwaiter().GetResult();

    // ─── IPdfReportService ───
    public byte[] GenerateExecutivePdf(DatosNacionalesDto datos)
        => CallPythonApi("pdf-ejecutivo").GetAwaiter().GetResult();

    // ─── IExcelEnhancedService ───
    public byte[] GenerateEnhancedExcel(DatosNacionalesDto datos)
        => CallPythonApi("excel-enhanced").GetAwaiter().GetResult();

    // ─── IWordOperatividadService ───
    public byte[] GenerateOperatividadDocx(DatosNacionalesDto datos)
        => CallPythonApi("word-operatividad").GetAwaiter().GetResult();

    // ─── IPptReportService ───
    public byte[] GeneratePptDinamico(DatosNacionalesDto datos, PptFilterDto? filtros = null)
    {
        var extraParams = new Dictionary<string, string>();
        if (filtros != null)
        {
            if (!string.IsNullOrEmpty(filtros.Empresa))
                extraParams["empresa"] = filtros.Empresa;
            if (filtros.Departamentos.Any())
                extraParams["departamentos"] = string.Join(",", filtros.Departamentos);
            if (filtros.Provincias.Any())
                extraParams["provincias"] = string.Join(",", filtros.Provincias);
            if (filtros.Distritos.Any())
                extraParams["distritos"] = string.Join(",", filtros.Distritos);
            if (filtros.TiposSiniestro.Any())
                extraParams["tipos_siniestro"] = string.Join(",", filtros.TiposSiniestro);
            if (filtros.FechaInicio.HasValue)
                extraParams["fecha_inicio"] = filtros.FechaInicio.Value.ToString("yyyy-MM-dd");
            if (filtros.FechaFin.HasValue)
                extraParams["fecha_fin"] = filtros.FechaFin.Value.ToString("yyyy-MM-dd");
        }
        return CallPythonApi("ppt-dinamico", null, extraParams).GetAwaiter().GetResult();
    }

    public byte[] GeneratePptHistorico(DatosNacionalesDto datos, string departamento)
        => CallPythonApi("ppt-historico", departamento).GetAwaiter().GetResult();

    // ─── Core HTTP call ───
    private async Task<byte[]> CallPythonApi(string reportType, string? departamento = null, Dictionary<string, string>? extraParams = null)
    {
        if (!_fileStore.HasFiles)
            throw new InvalidOperationException("No hay archivos Excel cargados. Suba los archivos primero.");

        using var content = new MultipartFormDataContent();

        // Add Excel files from scoped store (per-user)
        var midagriContent = new ByteArrayContent(_fileStore.MidagriBytes!);
        midagriContent.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        content.Add(midagriContent, "midagri", "midagri.xlsx");

        var siniestrosContent = new ByteArrayContent(_fileStore.SiniestrosBytes!);
        siniestrosContent.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        content.Add(siniestrosContent, "siniestros", "siniestros.xlsx");

        // Build URL with query params
        var url = $"{_baseUrl}/api/process-and-generate?report_type={reportType}";
        if (!string.IsNullOrEmpty(departamento))
            url += $"&departamento={Uri.EscapeDataString(departamento)}";
        if (extraParams != null)
        {
            foreach (var kv in extraParams)
                url += $"&{kv.Key}={Uri.EscapeDataString(kv.Value)}";
        }

        var response = await _httpClient.PostAsync(url, content);

        if (!response.IsSuccessStatusCode)
        {
            var errorBody = await response.Content.ReadAsStringAsync();
            throw new InvalidOperationException($"Error del servicio de reportes ({response.StatusCode}): {errorBody}");
        }

        return await response.Content.ReadAsByteArrayAsync();
    }
}
