namespace Infrastructure;

using Application.Contracts.Repositories;
using Application.Contracts.Services;
using Infrastructure.Alertas;
using Infrastructure.ExcelReader;
using Infrastructure.PythonApi;
using Infrastructure.StaticData;
using Microsoft.Extensions.DependencyInjection;

public static class DependencyInjection
{
    public static IServiceCollection AddInfrastructure(this IServiceCollection services)
    {
        // Data access (keep in .NET)
        services.AddScoped<IExcelReaderRepository, ExcelReaderRepository>();
        services.AddScoped<IStaticDataRepository, StaticDataRepository>();

        // Business logic (keep in .NET)
        services.AddScoped<ISemaforoService, SemaforoService>();

        // Document generation (delegate to Python API)
        services.AddHttpClient<PythonApiReportService>();
        services.AddScoped<IWordReportService>(sp => sp.GetRequiredService<PythonApiReportService>());
        services.AddScoped<IExcelReportService>(sp => sp.GetRequiredService<PythonApiReportService>());
        services.AddScoped<IPdfReportService>(sp => sp.GetRequiredService<PythonApiReportService>());
        services.AddScoped<IExcelEnhancedService>(sp => sp.GetRequiredService<PythonApiReportService>());
        services.AddScoped<IWordOperatividadService>(sp => sp.GetRequiredService<PythonApiReportService>());
        services.AddScoped<IPptReportService>(sp => sp.GetRequiredService<PythonApiReportService>());

        return services;
    }
}
