namespace Application.Contracts.Services;

using Application.DTOs.Responses;

public interface IWordOperatividadService
{
    byte[] GenerateOperatividadDocx(DatosNacionalesDto datos);
}
