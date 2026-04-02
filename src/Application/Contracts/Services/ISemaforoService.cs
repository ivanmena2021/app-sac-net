namespace Application.Contracts.Services;

using Application.DTOs.Responses;

public interface ISemaforoService
{
    SemaforoResultDto ComputeSemaforo(DatosNacionalesDto datos);
}
