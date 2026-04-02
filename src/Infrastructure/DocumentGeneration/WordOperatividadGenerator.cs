namespace Infrastructure.DocumentGeneration;

using Application.Contracts.Services;
using Application.DTOs.Responses;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

public class WordOperatividadGenerator : IWordOperatividadService
{
    private static readonly CultureInfo EsPe = new("es-PE");

    private static string Fmt(double val) => val.ToString("N2", EsPe);
    private static string FmtInt(double val) => ((int)val).ToString("N0", EsPe);
    private static string FmtPct(double val) => $"{val:F2}%";

    public byte[] GenerateOperatividadDocx(DatosNacionalesDto datos)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            var sectionProps = new SectionProperties(
                new PageMargin { Top = 1008, Right = 1224, Bottom = 864, Left = 1224 }
            );

            // === HEADER ===
            AddCenteredParagraph(body, "AYUDA MEMORIA OPERATIVIDAD SAC", 13, true, "1F4E79");
            AddCenteredParagraph(body, "CAMPAÑA AGRÍCOLA 2025-2026", 11, true, "2E75B6");
            AddCenteredParagraph(body, $"(AL {datos.FechaCorte})", 10, false, "666666");
            AddEmptyParagraph(body);

            var m = datos.Metricas;

            // === SECTION: OPERATIVIDAD ===
            AddHeading(body, "I. OPERATIVIDAD", "1F4E79");

            // a) Avisos de siniestros
            AddSubHeading(body, "a) Avisos de siniestros");
            AddBullet(body, $"Total de avisos reportados: {FmtInt(m.TotalAvisos)}");

            // Top departments by count
            var deptCounts = datos.Midagri
                .GroupBy(s => s.Departamento)
                .Select(g => new { Dept = g.Key, Count = g.Count() })
                .OrderByDescending(x => x.Count)
                .Take(5)
                .ToList();

            if (deptCounts.Count > 0)
            {
                var topText = string.Join(", ", deptCounts.Select(d =>
                    $"{CultureInfo.CurrentCulture.TextInfo.ToTitleCase(d.Dept.ToLower())} ({d.Count})"));
                AddBullet(body, $"Departamentos con más avisos: {topText}");
            }

            // Top disaster types
            var topTipos = datos.SiniestrosPorTipo.OrderByDescending(x => x.Value).Take(3).ToList();
            if (topTipos.Count > 0)
            {
                var topText = string.Join(", ", topTipos.Select(t =>
                    $"{CultureInfo.CurrentCulture.TextInfo.ToTitleCase(t.Key.ToLower())} ({t.Value})"));
                AddBullet(body, $"Tipos de siniestro predominantes: {topText}");
            }

            AddEmptyParagraph(body);

            // b) Resultados
            AddSubHeading(body, "b) Resultados de evaluación");
            AddBullet(body, $"Avisos ajustados (cerrados): {FmtInt(m.TotalAjustados)} ({m.PctAjustados:F1}%)");
            AddBullet(body, $"Indemnizaciones reconocidas: S/ {Fmt(m.MontoIndemnizado)}");
            AddBullet(body, $"Índice de siniestralidad general: {FmtPct(m.IndiceSiniestralidad)}");
            AddEmptyParagraph(body);

            // === CUADRO: Siniestralidad por Departamento ===
            AddHeading(body, "Cuadro: Indemnizaciones y Siniestralidad por Departamento", "1F4E79");

            var headers = new[] { "Departamento", "Ha Indemn.", "Indemnización (S/)", "Prima Neta (S/)", "Siniestralidad" };
            var rows = new List<string[]>();

            // Build from cuadro2 + materia
            var materiaLookup = datos.Materia.ToDictionary(m2 => m2.Departamento, m2 => m2, StringComparer.OrdinalIgnoreCase);
            foreach (var r in datos.Cuadro2.Where(r => r.Departamento != "TOTAL").OrderBy(r => r.Departamento))
            {
                materiaLookup.TryGetValue(r.Departamento, out var mat);
                var primaNeta = mat?.PrimaNeta ?? 0;
                var siniestralidad = primaNeta > 0 ? r.MontoIndemnizado / primaNeta * 100 : 0;
                rows.Add(new[] {
                    CultureInfo.CurrentCulture.TextInfo.ToTitleCase(r.Departamento.ToLower()),
                    Fmt(r.HaIndemnizadas),
                    Fmt(r.MontoIndemnizado),
                    Fmt(primaNeta),
                    FmtPct(siniestralidad)
                });
            }

            // Total row
            var totalPrima = datos.Materia.Sum(m2 => m2.PrimaNeta);
            var totalSin = totalPrima > 0 ? m.MontoIndemnizado / totalPrima * 100 : 0;
            rows.Add(new[] { "TOTAL", Fmt(m.HaIndemnizadas), Fmt(m.MontoIndemnizado), Fmt(totalPrima), FmtPct(totalSin) });

            AddTable(body, headers, rows, "1F4E79");
            AddEmptyParagraph(body);

            // === SECTION: DESEMBOLSOS ===
            AddHeading(body, "II. DESEMBOLSOS", "1F4E79");
            AddBullet(body, $"Monto total desembolsado: S/ {Fmt(m.MontoDesembolsado)}");
            AddBullet(body, $"Avance de desembolso: {m.PctDesembolso:F1}%");
            AddBullet(body, $"Productores beneficiados: {FmtInt(m.ProductoresDesembolso)}");
            AddEmptyParagraph(body);

            // Desembolso table
            AddHeading(body, "Cuadro: Desembolsos por Departamento", "1F4E79");
            var dHeaders = new[] { "Departamento", "Indemnización (S/)", "Desembolso (S/)", "% Desembolso", "Productores" };
            var dRows = datos.Cuadro2.Where(r => r.Departamento != "TOTAL" && r.MontoIndemnizado > 0)
                .OrderByDescending(r => r.MontoDesembolsado)
                .Select(r => new[] {
                    CultureInfo.CurrentCulture.TextInfo.ToTitleCase(r.Departamento.ToLower()),
                    Fmt(r.MontoIndemnizado),
                    Fmt(r.MontoDesembolsado),
                    FmtPct(r.MontoIndemnizado > 0 ? r.MontoDesembolsado / r.MontoIndemnizado * 100 : 0),
                    FmtInt(r.Productores)
                }).ToList();

            dRows.Add(new[] { "TOTAL", Fmt(m.MontoIndemnizado), Fmt(m.MontoDesembolsado), FmtPct(m.PctDesembolso), FmtInt(m.ProductoresDesembolso) });
            AddTable(body, dHeaders, dRows, "1F4E79");

            // Nota final
            AddEmptyParagraph(body);
            var nota = new Paragraph();
            nota.Append(CreateRun("Nota: ", 9, true, "000000", true));
            nota.Append(CreateRun("La vigencia de la póliza es del 01/08/2025 al 01/08/2026.", 9, false, "000000", true));
            body.Append(nota);

            body.Append(sectionProps);
        }

        return ms.ToArray();
    }

    // === HELPER METHODS (same pattern as WordDocumentGenerator) ===

    private static Run CreateRun(string text, int fontSize, bool bold, string color, bool italic = false)
    {
        var run = new Run();
        var props = new RunProperties();
        props.Append(new RunFonts { Ascii = "Arial Narrow", HighAnsi = "Arial Narrow" });
        props.Append(new FontSize { Val = (fontSize * 2).ToString() });
        if (bold) props.Append(new Bold());
        if (italic) props.Append(new Italic());
        props.Append(new Color { Val = color });
        run.Append(props);
        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    private static void AddCenteredParagraph(Body body, string text, int fontSize, bool bold, string color)
    {
        var para = new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }));
        para.Append(CreateRun(text, fontSize, bold, color));
        body.Append(para);
    }

    private static void AddHeading(Body body, string text, string color)
    {
        var para = new Paragraph(new ParagraphProperties(new SpacingBetweenLines { Before = "240", After = "120" }));
        para.Append(CreateRun(text, 12, true, color));
        body.Append(para);
    }

    private static void AddSubHeading(Body body, string text)
    {
        var para = new Paragraph(new ParagraphProperties(new SpacingBetweenLines { Before = "160", After = "80" }));
        para.Append(CreateRun(text, 11, true, "2E75B6"));
        body.Append(para);
    }

    private static void AddBullet(Body body, string text)
    {
        var para = new Paragraph();
        para.Append(CreateRun($"• {text}", 10, false, "000000"));
        body.Append(para);
    }

    private static void AddEmptyParagraph(Body body) => body.Append(new Paragraph());

    private static void AddTable(Body body, string[] headers, List<string[]> rows, string headerBgColor)
    {
        var table = new Table();
        var tblProps = new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4, Color = "AAAAAA" },
                new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "AAAAAA" },
                new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "AAAAAA" },
                new RightBorder { Val = BorderValues.Single, Size = 4, Color = "AAAAAA" },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "AAAAAA" },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "AAAAAA" }
            ),
            new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }
        );
        table.Append(tblProps);

        // Header row
        var headerRow = new TableRow();
        foreach (var h in headers)
        {
            var cell = new TableCell(
                new TableCellProperties(new Shading { Val = ShadingPatternValues.Clear, Fill = headerBgColor }),
                new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }))
                { }
            );
            cell.Elements<Paragraph>().First().Append(CreateRun(h, 8, true, "FFFFFF"));
            headerRow.Append(cell);
        }
        table.Append(headerRow);

        // Data rows
        for (int i = 0; i < rows.Count; i++)
        {
            var isTotal = rows[i][0].ToUpper() == "TOTAL";
            var isAlt = i % 2 == 0;
            var dataRow = new TableRow();

            for (int j = 0; j < rows[i].Length; j++)
            {
                var cell = new TableCell();
                var cellProps = new TableCellProperties();
                if (isTotal) cellProps.Append(new Shading { Val = ShadingPatternValues.Clear, Fill = headerBgColor });
                else if (isAlt) cellProps.Append(new Shading { Val = ShadingPatternValues.Clear, Fill = "F2F7FB" });
                cell.Append(cellProps);

                var align = j > 0 ? JustificationValues.Right : JustificationValues.Left;
                var para = new Paragraph(new ParagraphProperties(new Justification { Val = align }));
                para.Append(CreateRun(rows[i][j], 8, isTotal, isTotal ? "FFFFFF" : "000000"));
                cell.Append(para);
                dataRow.Append(cell);
            }
            table.Append(dataRow);
        }
        body.Append(table);
    }
}
