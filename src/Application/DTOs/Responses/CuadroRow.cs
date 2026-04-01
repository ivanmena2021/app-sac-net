namespace Application.DTOs.Responses;

public class Cuadro1Row
{
    public string Departamento { get; set; } = string.Empty;
    public double PrimaTotal { get; set; }
    public double Hectareas { get; set; }
    public double SumaAsegurada { get; set; }
}

public class Cuadro2Row
{
    public string Departamento { get; set; } = string.Empty;
    public double HaIndemnizadas { get; set; }
    public double MontoIndemnizado { get; set; }
    public double MontoDesembolsado { get; set; }
    public double Productores { get; set; }
}

public class Cuadro3Row
{
    public string Departamento { get; set; } = string.Empty;
    public int Avisos { get; set; }
    public double HaIndemn { get; set; }
    public double MontoIndemnizado { get; set; }
    public double MontoDesembolsado { get; set; }
    public double Productores { get; set; }
}
