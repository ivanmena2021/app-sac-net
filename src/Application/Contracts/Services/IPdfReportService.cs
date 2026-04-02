namespace Application.Contracts.Services;

using Application.DTOs.Responses;

public interface IPdfReportService
{
    byte[] GenerateExecutivePdf(DatosNacionalesDto datos);
}
