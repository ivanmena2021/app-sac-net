namespace Application.Contracts.Repositories;

using Domain.Entities;

public interface IExcelReaderRepository
{
    List<Siniestro> ReadMidagriExcel(Stream fileStream);
    List<Siniestro> ReadSiniestrosExcel(Stream fileStream);
}
