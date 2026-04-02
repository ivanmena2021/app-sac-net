namespace Application.Contracts.Services;

using Application.DTOs.Responses;

public interface IExcelEnhancedService
{
    byte[] GenerateEnhancedExcel(DatosNacionalesDto datos);
}
