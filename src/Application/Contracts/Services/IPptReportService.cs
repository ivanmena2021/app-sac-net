namespace Application.Contracts.Services;

using Application.DTOs.Requests;
using Application.DTOs.Responses;

public interface IPptReportService
{
    byte[] GeneratePptDinamico(DatosNacionalesDto datos, PptFilterDto? filtros = null);
    byte[] GeneratePptHistorico(DatosNacionalesDto datos, string departamento);
}
