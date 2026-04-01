namespace Application.DTOs.Responses;

using Domain.Entities;

public class DatosNacionalesDto
{
    public string FechaCorte { get; set; } = string.Empty;
    public MetricasDto Metricas { get; set; } = new();
    public string EmpresasText { get; set; } = string.Empty;
    public Dictionary<string, int> Empresas { get; set; } = new();

    // Cuadros
    public List<Cuadro1Row> Cuadro1 { get; set; } = new();
    public List<Cuadro2Row> Cuadro2 { get; set; } = new();
    public List<Cuadro3Row> Cuadro3 { get; set; } = new();

    // Lluvias
    public int TotalLluvia { get; set; }
    public double PctLluvia { get; set; }
    public Dictionary<string, int> LluviaPorTipo { get; set; } = new();

    // Siniestros
    public Dictionary<string, int> SiniestrosPorTipo { get; set; } = new();
    public Dictionary<string, int> Top3Siniestros { get; set; } = new();

    // Departamentos
    public List<string> DepartamentosList { get; set; } = new();

    // Raw data (for EME report and departmental details)
    public List<Siniestro> Midagri { get; set; } = new();
    public List<Siniestro> SiniestrosOriginal { get; set; } = new();
    public List<MateriaAsegurada> Materia { get; set; } = new();
}
