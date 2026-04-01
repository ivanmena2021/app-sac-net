namespace Domain.Entities;

using Domain.Common;

public class MateriaAsegurada : BaseEntity
{
    public string Departamento { get; set; } = string.Empty;
    public string Capital { get; set; } = string.Empty;
    public string EmpresaAseguradora { get; set; } = string.Empty;
    public string CultivosAsegurados { get; set; } = string.Empty;
    public double PrimaTotal { get; set; }
    public double PrimaNeta { get; set; }
    public double SuperficieAsegurada { get; set; }
    public double ProductoresAsegurados { get; set; }
    public double ValoresAsegurados { get; set; }
    public string Disparador { get; set; } = string.Empty;
    public double SumaAseguradaHa { get; set; }
}
