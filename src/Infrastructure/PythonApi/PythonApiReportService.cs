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
/// </summary>
public class PythonApiReportService : IWordReportService, IExcelReportService,
    IPdfReportService, IExcelEnhancedService, IWordOperatividadService, IPptReportService
{
    private readonly HttpClient _httpClient;
    private readonly string _baseUrl;

    // Store uploaded file bytes for reuse across generate calls
    private static byte[]? _lastMidagriBytes;
    private static byte[]? _lastSiniestrosBytes;

    public PythonApiReportService(HttpClient httpClient, IConfiguration configuration)
    {
        _httpClient = httpClient;
        _baseUrl = configuration["PythonApi:BaseUrl"] ?? "http://localhost:8000";
        _httpClient.Timeout = TimeSpan.FromMinutes(5); // PPT generation can be slow
    }

    /// <summary>
    /// Store the uploaded Excel files for later use by generate methods.
    /// Called from Home.razor after file upload.
    /// </summary>
    public static void SetUploadedFiles(byte[] midagriBytes, byte[] siniestrosBytes)
    {
        _lastMidagriBytes = midagriBytes;
        _lastSiniestrosBytes = siniestrosBytes;
    }

    public static void ClearUploadedFiles()
    {
        _lastMidagriBytes = null;
        _lastSiniestrosBytes = null;
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
        string? depto = filtros?.Departamentos.FirstOrDefault();
        return CallPythonApi("ppt-dinamico", depto).GetAwaiter().GetResult();
    }

    public byte[] GeneratePptHistorico(DatosNacionalesDto datos, string departamento)
        => CallPythonApi("ppt-historico", departamento).GetAwaiter().GetResult();

    // ─── Core HTTP call ───
    private async Task<byte[]> CallPythonApi(string reportType, string? departamento = null)
    {
        if (_lastMidagriBytes == null || _lastSiniestrosBytes == null)
            throw new InvalidOperationException("No hay archivos Excel cargados. Suba los archivos primero.");

        using var content = new MultipartFormDataContent();

        // Add Excel files
        var midagriContent = new ByteArrayContent(_lastMidagriBytes);
        midagriContent.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        content.Add(midagriContent, "midagri", "midagri.xlsx");

        var siniestrosContent = new ByteArrayContent(_lastSiniestrosBytes);
        siniestrosContent.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        content.Add(siniestrosContent, "siniestros", "siniestros.xlsx");

        // Build URL with query params
        var url = $"{_baseUrl}/api/process-and-generate?report_type={reportType}";
        if (!string.IsNullOrEmpty(departamento))
            url += $"&departamento={Uri.EscapeDataString(departamento)}";

        var response = await _httpClient.PostAsync(url, content);

        if (!response.IsSuccessStatusCode)
        {
            var errorBody = await response.Content.ReadAsStringAsync();
            throw new InvalidOperationException($"Error del servicio de reportes ({response.StatusCode}): {errorBody}");
        }

        return await response.Content.ReadAsByteArrayAsync();
    }
}
