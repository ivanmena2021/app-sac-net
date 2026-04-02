namespace Infrastructure;

using Application.Contracts.Repositories;
using Application.Contracts.Services;
using Infrastructure.Alertas;
using Infrastructure.DocumentGeneration;
using Infrastructure.ExcelReader;
using Infrastructure.StaticData;
using Microsoft.Extensions.DependencyInjection;

public static class DependencyInjection
{
    public static IServiceCollection AddInfrastructure(this IServiceCollection services)
    {
        services.AddScoped<IExcelReaderRepository, ExcelReaderRepository>();
        services.AddScoped<IStaticDataRepository, StaticDataRepository>();
        services.AddScoped<IWordReportService, WordDocumentGenerator>();
        services.AddScoped<IExcelReportService, ExcelDocumentGenerator>();
        services.AddScoped<IPdfReportService, PdfReportGenerator>();
        services.AddScoped<IExcelEnhancedService, ExcelEnhancedGenerator>();
        services.AddScoped<IWordOperatividadService, WordOperatividadGenerator>();
        services.AddScoped<ISemaforoService, SemaforoService>();
        return services;
    }
}
