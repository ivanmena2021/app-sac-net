namespace Application.Contracts.Repositories;

using Application.DTOs.Responses;
using Domain.Entities;

public interface IStaticDataRepository
{
    List<MateriaAsegurada> LoadMateriaAsegurada();

    /// <summary>resumen_departamental.json → {depto: {campaña: datos}}</summary>
    Dictionary<string, Dictionary<string, DeptoCampanaDto>> LoadResumenDepartamental();

    /// <summary>resumen_campanas.json → {campaña: resumen}</summary>
    Dictionary<string, object> LoadResumenCampanas();

    /// <summary>Primas_Totales_SAC_2020-2026.xlsx → {campaña: {depto: prima_neta}}</summary>
    Dictionary<string, Dictionary<string, double>> LoadPrimasHistoricas();

    /// <summary>series_temporales.json</summary>
    SeriesTemporalesDto LoadSeriesTemporales();

    /// <summary>calendario_cultivos_historico.json → {depto: List cultivos}</summary>
    Dictionary<string, List<CalendarioCultivoDto>> LoadCalendarioAgricola();
}
