namespace Application.DTOs.Responses;

public class MetricasDto
{
    public int TotalAvisos { get; set; }
    public int TotalAjustados { get; set; }
    public double PctAjustados { get; set; }
    public double HaIndemnizadas { get; set; }
    public double MontoIndemnizado { get; set; }
    public double MontoDesembolsado { get; set; }
    public int ProductoresDesembolso { get; set; }
    public double PrimaTotal { get; set; }
    public double PrimaNeta { get; set; }
    public double SupAsegurada { get; set; }
    public int ProdAsegurados { get; set; }
    public double IndiceSiniestralidad { get; set; }
    public double PctDesembolso { get; set; }
    public int DeptosConDesembolso { get; set; }
}
