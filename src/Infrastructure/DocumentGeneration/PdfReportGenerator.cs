namespace Infrastructure.DocumentGeneration;

using Application.Contracts.Services;
using Application.DTOs.Responses;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System.Globalization;

public class PdfReportGenerator : IPdfReportService
{
    private static readonly CultureInfo EsPe = new("es-PE");

    static PdfReportGenerator()
    {
        QuestPDF.Settings.License = LicenseType.Community;
    }

    public byte[] GenerateExecutivePdf(DatosNacionalesDto datos)
    {
        var m = datos.Metricas;
        var document = Document.Create(container =>
        {
            container.Page(page =>
            {
                page.Size(PageSizes.A4);
                page.MarginHorizontal(20, Unit.Millimetre);
                page.MarginVertical(15, Unit.Millimetre);

                page.Header().Height(32, Unit.Millimetre).Background("#0c2340").Padding(6, Unit.Millimetre).Column(col =>
                {
                    col.Item().Text("SEGURO AGRÍCOLA CATASTRÓFICO").FontSize(16).Bold().FontColor("#FFFFFF");
                    col.Item().Text($"Resumen Ejecutivo | SAC 2025-2026 | Corte al {datos.FechaCorte}").FontSize(10).FontColor("#87CEEB");
                });

                page.Content().Column(content =>
                {
                    // KPI Boxes - 2 rows of 4
                    content.Item().PaddingTop(8, Unit.Millimetre).Row(row =>
                    {
                        AddKpiBox(row, "Total Avisos", m.TotalAvisos.ToString("N0", EsPe));
                        AddKpiBox(row, "Ha Indemnizadas", m.HaIndemnizadas.ToString("N2", EsPe));
                        AddKpiBox(row, "Monto Indemnizado", $"S/ {m.MontoIndemnizado.ToString("N0", EsPe)}");
                        AddKpiBox(row, "Monto Desembolsado", $"S/ {m.MontoDesembolsado.ToString("N0", EsPe)}");
                    });
                    content.Item().PaddingTop(2, Unit.Millimetre).Row(row =>
                    {
                        AddKpiBox(row, "Productores", m.ProductoresDesembolso.ToString("N0", EsPe));
                        AddKpiBox(row, $"Siniestralidad", $"{m.IndiceSiniestralidad:F1}%");
                        AddKpiBox(row, "% Desembolso", $"{m.PctDesembolso:F1}%");
                        AddKpiBox(row, "Prima Total", $"S/ {m.PrimaTotal.ToString("N0", EsPe)}");
                    });

                    // Top 10 Departments Table
                    content.Item().PaddingTop(8, Unit.Millimetre).Text("Top 10 Departamentos por Indemnización").FontSize(11).Bold().FontColor("#0c2340");
                    content.Item().PaddingTop(4, Unit.Millimetre).Table(table =>
                    {
                        table.ColumnsDefinition(cols =>
                        {
                            cols.RelativeColumn(3); // Departamento
                            cols.RelativeColumn(2); // Ha Indemn
                            cols.RelativeColumn(2); // Monto Indemn
                            cols.RelativeColumn(2); // Monto Desemb
                            cols.RelativeColumn(2); // Productores
                        });

                        // Header
                        var headers = new[] { "Departamento", "Ha Indemn.", "Monto Indemn.", "Monto Desemb.", "Productores" };
                        foreach (var h in headers)
                            table.Cell().Background("#1a5276").Padding(3).Text(h).FontSize(7).Bold().FontColor("#FFFFFF");

                        // Data rows (top 10 by MontoIndemnizado)
                        var top10 = datos.Cuadro2
                            .Where(r => r.Departamento != "TOTAL")
                            .OrderByDescending(r => r.MontoIndemnizado)
                            .Take(10)
                            .ToList();

                        for (int i = 0; i < top10.Count; i++)
                        {
                            var r = top10[i];
                            var bg = i % 2 == 0 ? "#FFFFFF" : "#F5F7FA";
                            table.Cell().Background(bg).Padding(2).Text(CultureInfo.CurrentCulture.TextInfo.ToTitleCase(r.Departamento.ToLower())).FontSize(7);
                            table.Cell().Background(bg).Padding(2).AlignRight().Text(r.HaIndemnizadas.ToString("N2", EsPe)).FontSize(7);
                            table.Cell().Background(bg).Padding(2).AlignRight().Text(r.MontoIndemnizado.ToString("N0", EsPe)).FontSize(7);
                            table.Cell().Background(bg).Padding(2).AlignRight().Text(r.MontoDesembolsado.ToString("N0", EsPe)).FontSize(7);
                            table.Cell().Background(bg).Padding(2).AlignRight().Text(r.Productores.ToString("N0", EsPe)).FontSize(7);
                        }
                    });

                    // Siniestros distribution
                    content.Item().PaddingTop(10, Unit.Millimetre).Background("#0c2340").Padding(4, Unit.Millimetre)
                        .Text("Distribución de Siniestros por Tipo").FontSize(10).Bold().FontColor("#FFFFFF");

                    content.Item().PaddingTop(4, Unit.Millimetre).Table(table =>
                    {
                        table.ColumnsDefinition(cols =>
                        {
                            cols.RelativeColumn(4);
                            cols.RelativeColumn(1);
                            cols.RelativeColumn(1);
                        });

                        table.Cell().Background("#1a5276").Padding(3).Text("Tipo de Siniestro").FontSize(8).Bold().FontColor("#FFFFFF");
                        table.Cell().Background("#1a5276").Padding(3).AlignRight().Text("Cantidad").FontSize(8).Bold().FontColor("#FFFFFF");
                        table.Cell().Background("#1a5276").Padding(3).AlignRight().Text("%").FontSize(8).Bold().FontColor("#FFFFFF");

                        var sorted = datos.SiniestrosPorTipo.OrderByDescending(x => x.Value).Take(10).ToList();
                        for (int i = 0; i < sorted.Count; i++)
                        {
                            var kvp = sorted[i];
                            var bg = i % 2 == 0 ? "#FFFFFF" : "#F5F7FA";
                            var pct = m.TotalAvisos > 0 ? (double)kvp.Value / m.TotalAvisos * 100 : 0;
                            table.Cell().Background(bg).Padding(2).Text(CultureInfo.CurrentCulture.TextInfo.ToTitleCase(kvp.Key.ToLower())).FontSize(7);
                            table.Cell().Background(bg).Padding(2).AlignRight().Text(kvp.Value.ToString("N0")).FontSize(7);
                            table.Cell().Background(bg).Padding(2).AlignRight().Text($"{pct:F1}%").FontSize(7);
                        }
                    });

                    // Observations
                    content.Item().PaddingTop(8, Unit.Millimetre).Background("#0c2340").Padding(4, Unit.Millimetre)
                        .Text("Observaciones Clave").FontSize(10).Bold().FontColor("#FFFFFF");

                    var obs = GenerateObservations(datos);
                    for (int i = 0; i < obs.Count; i++)
                    {
                        content.Item().PaddingTop(2, Unit.Millimetre).Row(row =>
                        {
                            row.ConstantItem(15).Text($"{i + 1}.").FontSize(8).Bold().FontColor("#1a5276");
                            row.RelativeItem().Text(obs[i]).FontSize(8).FontColor("#333333");
                        });
                    }

                    // Disclaimer
                    content.Item().PaddingTop(10, Unit.Millimetre).Text(
                        "Documento generado automáticamente. Los datos corresponden al corte indicado y pueden variar."
                    ).FontSize(6.5f).Italic().FontColor("#999999");
                });

                page.Footer().AlignCenter().Text(text =>
                {
                    text.Span($"SAC 2025-2026 | Generado: {DateTime.Now:dd/MM/yyyy HH:mm} | Página ").FontSize(7).FontColor("#999999");
                    text.CurrentPageNumber().FontSize(7).FontColor("#999999");
                    text.Span(" / ").FontSize(7).FontColor("#999999");
                    text.TotalPages().FontSize(7).FontColor("#999999");
                });
            });
        });

        using var ms = new MemoryStream();
        document.GeneratePdf(ms);
        return ms.ToArray();
    }

    private static void AddKpiBox(RowDescriptor row, string label, string value)
    {
        row.RelativeItem().Border(0.5f).BorderColor("#E0E0E0").Background("#F8F9FA").Padding(4, Unit.Millimetre).Column(col =>
        {
            col.Item().Text(label).FontSize(6.5f).FontColor("#7f8c8d");
            col.Item().PaddingTop(2, Unit.Millimetre).Text(value).FontSize(10).Bold().FontColor("#0c2340");
        });
    }

    private static List<string> GenerateObservations(DatosNacionalesDto datos)
    {
        var obs = new List<string>();
        var m = datos.Metricas;

        if (m.TotalAvisos > 0)
            obs.Add($"Se han registrado {m.TotalAvisos:N0} avisos de siniestro en {datos.DepartamentosList.Count} departamentos.");

        if (m.IndiceSiniestralidad > 100)
            obs.Add($"El índice de siniestralidad es {m.IndiceSiniestralidad:F1}%, superando la prima neta recaudada.");
        else if (m.IndiceSiniestralidad > 70)
            obs.Add($"El índice de siniestralidad es {m.IndiceSiniestralidad:F1}%, acercándose al total de la prima neta.");
        else if (m.IndiceSiniestralidad > 0)
            obs.Add($"El índice de siniestralidad se encuentra en {m.IndiceSiniestralidad:F1}%.");

        if (m.PctDesembolso > 0)
        {
            var pendiente = m.MontoIndemnizado - m.MontoDesembolsado;
            obs.Add($"Se ha desembolsado el {m.PctDesembolso:F1}% del monto indemnizado. Pendiente: S/ {pendiente:N0}.");
        }

        if (m.ProductoresDesembolso > 0)
            obs.Add($"{m.ProductoresDesembolso:N0} productores han sido beneficiados con desembolsos.");

        if (datos.Top3Siniestros.Count > 0)
        {
            var top = string.Join(", ", datos.Top3Siniestros.Select(k =>
                $"{CultureInfo.CurrentCulture.TextInfo.ToTitleCase(k.Key.ToLower())} ({k.Value})"));
            obs.Add($"Los principales siniestros son: {top}.");
        }

        return obs;
    }
}
