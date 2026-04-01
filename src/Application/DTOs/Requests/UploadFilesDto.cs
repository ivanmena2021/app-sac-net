namespace Application.DTOs.Requests;

public class UploadFilesDto
{
    public Stream MidagriStream { get; set; } = Stream.Null;
    public Stream SiniestrosStream { get; set; } = Stream.Null;
    public string MidagriFileName { get; set; } = string.Empty;
    public string SiniestrosFileName { get; set; } = string.Empty;
}
