namespace Domain.Entities;

using Domain.Common;

public class Siniestro : BaseEntity
{
    public string Campana { get; set; } = string.Empty;
    public string CodigoAviso { get; set; } = string.Empty;
    public string Departamento { get; set; } = string.Empty;
    public string Provincia { get; set; } = string.Empty;
    public string Distrito { get; set; } = string.Empty;
    public string SectorEstadistico { get; set; } = string.Empty;
    public string TipoCultivo { get; set; } = string.Empty;
    public string Fenologia { get; set; } = string.Empty;
    public DateTime? FechaSiembra { get; set; }
    public DateTime? FechaCosecha { get; set; }
    public double SupSembrada { get; set; }
    public double SupAsegurada { get; set; }
    public string TipoSiniestro { get; set; } = string.Empty;
    public DateTime? FechaSiniestro { get; set; }
    public DateTime? FechaAviso { get; set; }
    public DateTime? FechaAtencion { get; set; }
    public string EstadoSiniestro { get; set; } = string.Empty;
    public string EstadoInspeccion { get; set; } = string.Empty;
    public double PrimaNetaDpto { get; set; }
    public string TipoCobertura { get; set; } = string.Empty;
    public double SupAfectada { get; set; }
    public double SupPerdida { get; set; }
    public string Dictamen { get; set; } = string.Empty;
    public double SupIndemnizada { get; set; }
    public double Indemnizacion { get; set; }
    public double MontoDesembolsado { get; set; }
    public double SupDesembolso { get; set; }
    public double NProductores { get; set; }
    public string CodigoPadron { get; set; } = string.Empty;
    public DateTime? FechaEnvioDras { get; set; }
    public DateTime? FechaValidacion { get; set; }
    public DateTime? FechaDesembolso { get; set; }
    public string Priorizado { get; set; } = string.Empty;
    public string Observacion { get; set; } = string.Empty;
}
