namespace Domain.Entities;

public class ResumenDepartamental
{
    public string Departamento { get; set; } = string.Empty;
    public int TotalAvisos { get; set; }
    public double HaIndemnizadas { get; set; }
    public double MontoIndemnizado { get; set; }
    public double MontoDesembolsado { get; set; }
    public double NProductores { get; set; }
    public double SupIndemnizada { get; set; }
    public double Indemnizacion { get; set; }
    public string Distritos { get; set; } = string.Empty;
}
