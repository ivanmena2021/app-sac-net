namespace Application.Contracts.Services;

public interface IAutoDownloadService
{
    bool IsConfigured();
    Task<AutoDownloadResult> DescargarRimacAsync(Action<string>? onProgress = null);
    Task<AutoDownloadResult> DescargarLaPositivaAsync(Action<string>? onProgress = null);
}

public class AutoDownloadResult
{
    public bool Success { get; set; }
    public byte[]? FileBytes { get; set; }
    public int Rows { get; set; }
    public string? FileName { get; set; }
    public string? Error { get; set; }
    public string Timestamp { get; set; } = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
}
