namespace Application.Contracts.Services;

using Application.DTOs.Responses;

public interface IDataProcessorService
{
    DatosNacionalesDto ProcessDynamicData(Stream midagriStream, Stream siniestrosStream);
    DatosDepartamentalDto GetDepartamentoData(DatosNacionalesDto datos, string departamento);
}
