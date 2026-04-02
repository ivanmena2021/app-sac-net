namespace Application.DTOs.Requests;

public class PptFilterDto
{
    public string? Empresa { get; set; }
    public List<string> Departamentos { get; set; } = new();
    public List<string> Provincias { get; set; } = new();
    public List<string> Distritos { get; set; } = new();
    public List<string> TiposSiniestro { get; set; } = new();
    public bool IncluirNacional { get; set; } = true;
    public bool FiltrarFecha { get; set; }
    public DateTime? FechaInicio { get; set; }
    public DateTime? FechaFin { get; set; }
}
