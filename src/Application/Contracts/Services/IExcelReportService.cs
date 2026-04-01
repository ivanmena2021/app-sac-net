namespace Application.Contracts.Services;

using Application.DTOs.Responses;

public interface IExcelReportService
{
    byte[] GenerateReporteEme(DatosNacionalesDto datos);
}
