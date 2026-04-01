namespace Application.Contracts.Services;

using Application.DTOs.Responses;

public interface IWordReportService
{
    byte[] GenerateNacionalDocx(DatosNacionalesDto datos);
    byte[] GenerateDepartamentalDocx(DatosDepartamentalDto datos);
}
