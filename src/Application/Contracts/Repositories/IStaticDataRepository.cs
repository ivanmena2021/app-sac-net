namespace Application.Contracts.Repositories;

using Domain.Entities;

public interface IStaticDataRepository
{
    List<MateriaAsegurada> LoadMateriaAsegurada();
}
