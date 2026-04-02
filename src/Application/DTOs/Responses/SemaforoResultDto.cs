namespace Application.DTOs.Responses;

public class SemaforoResultDto
{
    public int TotalAvisos { get; set; }
    public int Verde { get; set; }
    public int Ambar { get; set; }
    public int Rojo { get; set; }
    public List<SemaforoRow> Rows { get; set; } = new();
    public Dictionary<string, SemaforoDepartamento> PorDepartamento { get; set; } = new();
}

public class SemaforoRow
{
    public string CodigoAviso { get; set; } = "";
    public string Departamento { get; set; } = "";
    public string Provincia { get; set; } = "";
    public string Distrito { get; set; } = "";
    public string Etapa { get; set; } = "";
    public string Alerta { get; set; } = ""; // verde, ambar, rojo
    public int Dias { get; set; }
    public string Detalle { get; set; } = "";
}

public class SemaforoDepartamento
{
    public int Total { get; set; }
    public int Verde { get; set; }
    public int Ambar { get; set; }
    public int Rojo { get; set; }
}
