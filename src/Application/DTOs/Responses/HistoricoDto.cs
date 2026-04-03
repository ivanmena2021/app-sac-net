namespace Application.DTOs.Responses;

/// <summary>Métricas agregadas de una campaña (histórica o actual).</summary>
public class CampanaMetricsDto
{
    public string Campana { get; set; } = string.Empty;
    public int Avisos { get; set; }
    public int Indemnizados { get; set; }
    public double Monto { get; set; }
    public double Ha { get; set; }
    public double Desembolso { get; set; }
    public double PrimaNeta { get; set; }
    public double Siniestralidad { get; set; }
}

/// <summary>Datos departamentales por campaña (desde resumen_departamental.json).</summary>
public class DeptoCampanaDto
{
    public int Avisos { get; set; }
    public int Indemnizados { get; set; }
    public int PerdidaTotal { get; set; }
    public double MontoIndemnizado { get; set; }
    public double HaIndemnizadas { get; set; }
    public double MontoDesembolsado { get; set; }
    public int Provincias { get; set; }
    public int Distritos { get; set; }
}

/// <summary>Serie temporal mensual: avisos e indemnizaciones por campaña.</summary>
public class SeriesTemporalesDto
{
    /// <summary>{campaña: {periodo: count}}</summary>
    public Dictionary<string, Dictionary<string, int>> Avisos { get; set; } = new();
    /// <summary>{campaña: {periodo: {n, monto}}}</summary>
    public Dictionary<string, Dictionary<string, IndemnMensualDto>> Indemnizaciones { get; set; } = new();
}

public class IndemnMensualDto
{
    public int N { get; set; }
    public double Monto { get; set; }
}

/// <summary>Entrada del calendario agrícola por cultivo.</summary>
public class CalendarioCultivoDto
{
    public string Cultivo { get; set; } = string.Empty;
    public List<int> MesesSiembra { get; set; } = new();
    public List<int> MesesCosecha { get; set; } = new();
    public List<int> MesesRiesgo { get; set; } = new();
    public List<string> Riesgos { get; set; } = new();
    public List<string> GruposClimaticos { get; set; } = new();
    public int TotalAvisos { get; set; }
    public int TotalIndemnizados { get; set; }
    public int TotalPerdidaTotal { get; set; }
    public List<string> RiesgosAvisos { get; set; } = new();
    public List<string> RiesgosIndemnizados { get; set; } = new();
    public List<string> RiesgosPerdidaTotal { get; set; } = new();
}
