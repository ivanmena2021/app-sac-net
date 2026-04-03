namespace Application.DTOs.Responses;

public class DatosDepartamentalDto
{
    public string Departamento { get; set; } = string.Empty;
    public string Empresa { get; set; } = string.Empty;
    public double PrimaNeta { get; set; }
    public double PrimaTotal { get; set; }
    public double SupAsegurada { get; set; }
    public double HaAseguradas { get; set; }
    public int TotalAvisos { get; set; }
    public double HaIndemnizadas { get; set; }
    public double MontoIndemnizado { get; set; }
    public double MontoDesembolsado { get; set; }
    public int ProductoresDesembolso { get; set; }
    public int Indemnizables { get; set; }
    public int NoIndemnizables { get; set; }
    public string FechaCorte { get; set; } = string.Empty;
    public Dictionary<string, int> Estados { get; set; } = new();

    // For Word generation
    public List<string[]> AvisosTipo { get; set; } = new();
    public List<string[]> DistProvincia { get; set; } = new();
    public List<string> DistProvinciaHeaders { get; set; } = new()
    {
        "Provincia", "Avisos", "Sup. Indemn.", "Prod. Benef.", "Indemniz.", "Desembolso", "% Avance"
    };
    public List<string[]> EventosRecientes { get; set; } = new();
    public List<string> EventosHeaders { get; set; } = new()
    {
        "Fecha", "Provincia", "Distrito / Sector", "Cultivo", "Estado"
    };
    public string ResumenOperativo { get; set; } = string.Empty;
    public string ResumenDesembolso { get; set; } = string.Empty;
}
