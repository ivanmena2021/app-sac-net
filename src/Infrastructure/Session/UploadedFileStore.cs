namespace Infrastructure.Session;

using Application.Contracts.Services;

/// <summary>
/// Scoped implementation — one instance per Blazor Server circuit (user session).
/// Holds the uploaded Excel bytes so they can be forwarded to the Python API.
/// </summary>
public class UploadedFileStore : IUploadedFileStore
{
    public byte[]? MidagriBytes { get; set; }
    public byte[]? SiniestrosBytes { get; set; }
    public bool HasFiles => MidagriBytes != null && SiniestrosBytes != null;

    public void Clear()
    {
        MidagriBytes = null;
        SiniestrosBytes = null;
    }
}
