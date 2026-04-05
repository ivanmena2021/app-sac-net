namespace Application.Contracts.Services;

/// <summary>
/// Scoped service that stores uploaded Excel file bytes per-circuit (per-user session).
/// In Blazor Server, "Scoped" = one instance per SignalR circuit = per browser tab.
/// This replaces the previous static byte[] fields that were shared across all users.
/// </summary>
public interface IUploadedFileStore
{
    byte[]? MidagriBytes { get; set; }
    byte[]? SiniestrosBytes { get; set; }
    bool HasFiles => MidagriBytes != null && SiniestrosBytes != null;
    void Clear();
}
