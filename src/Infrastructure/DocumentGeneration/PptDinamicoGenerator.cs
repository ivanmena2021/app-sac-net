namespace Infrastructure.DocumentGeneration;

using Application.Contracts.Services;
using Application.DTOs.Requests;
using Application.DTOs.Responses;
using Domain.Entities;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using System.Globalization;

public class PptDinamicoGenerator : IPptReportService
{
    private static readonly CultureInfo EsPe = new("es-PE");

    public byte[] GeneratePptDinamico(DatosNacionalesDto datos, PptFilterDto? filtros = null)
    {
        var filtered = ApplyFilters(datos.Midagri, filtros);
        string scope = DetermineScope(filtros);
        string scopeDetail = BuildScopeDetail(filtros);

        int totalAvisos = filtered.Count;
        int totalAjustados = filtered.Count(s =>
            s.EstadoInspeccion?.Trim().Equals("CERRADO", StringComparison.OrdinalIgnoreCase) == true);
        double montoIndemn = filtered.Sum(s => s.Indemnizacion);
        double montoDesemb = filtered.Sum(s => s.MontoDesembolsado);
        double haIndemn = filtered.Sum(s => s.SupIndemnizada);
        int productores = (int)filtered.Where(s => s.Indemnizacion > 0).Sum(s => s.NProductores);
        double pctDesemb = montoIndemn > 0 ? montoDesemb / montoIndemn * 100 : 0;

        using var ms = new MemoryStream();
        using (var ppt = PresentationDocument.Create(ms, PresentationDocumentType.Presentation, true))
        {
            var presentationPart = ppt.AddPresentationPart();
            presentationPart.Presentation = new Presentation();
            presentationPart.Presentation.AppendChild(new SlideIdList());
            presentationPart.Presentation.AppendChild(
                new SlideSize { Cx = 12192000, Cy = 6858000 });
            presentationPart.Presentation.AppendChild(
                new NotesSize { Cx = 6858000, Cy = 9144000 });

            uint id = 256;

            // --- Slide 1: Title ---
            CreateTextSlide(presentationPart, ref id, new[]
            {
                ("SEGURO AGRÍCOLA CATASTRÓFICO — SAC 2025-2026", 24, true, "0C2340"),
                ($"Análisis {scope}", 16, false, "2980B9"),
                (scopeDetail, 12, false, "7F8C8D"),
                ($"Corte de datos: {datos.FechaCorte}", 12, false, "999999"),
                ("MIDAGRI — DGASFS / DSFFA", 10, false, "999999"),
            });

            // --- Slide 2: KPIs ---
            CreateTextSlide(presentationPart, ref id, new[]
            {
                ("Indicadores Principales", 22, true, "0C2340"),
                ($"Total Avisos: {totalAvisos.ToString("N0", EsPe)}", 14, true, "2980B9"),
                ($"Avisos Ajustados: {totalAjustados.ToString("N0", EsPe)} ({(totalAvisos > 0 ? (double)totalAjustados / totalAvisos * 100 : 0):F1}%)", 12, false, "333333"),
                ($"Hectáreas Indemnizadas: {haIndemn.ToString("N2", EsPe)}", 14, true, "27AE60"),
                ($"Indemnización Total: S/ {montoIndemn.ToString("N2", EsPe)}", 14, true, "E67E22"),
                ($"Desembolsos: S/ {montoDesemb.ToString("N2", EsPe)} ({pctDesemb:F1}%)", 12, false, "333333"),
                ($"Productores Beneficiados: {productores.ToString("N0", EsPe)}", 13, false, "27AE60"),
            });

            // --- Slide 3: Top Departments ---
            var deptTop = filtered
                .GroupBy(s => s.Departamento)
                .Select(g => (Dept: g.Key, Count: g.Count(), Indemn: g.Sum(s => s.Indemnizacion)))
                .OrderByDescending(x => x.Count)
                .Take(10)
                .ToList();

            var deptLines = new List<(string text, int fontSize, bool bold, string color)>
            {
                ("Top Departamentos por Avisos", 22, true, "0C2340")
            };
            foreach (var d in deptTop)
            {
                var name = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(d.Dept.ToLower());
                deptLines.Add(($"• {name}: {d.Count} avisos — S/ {d.Indemn.ToString("N0", EsPe)}", 11, false, "333333"));
            }
            CreateTextSlide(presentationPart, ref id, deptLines.ToArray());

            // --- Slide 4: Disaster Types ---
            var typeTop = filtered
                .GroupBy(s => s.TipoSiniestro)
                .Select(g => (Tipo: g.Key, Count: g.Count()))
                .OrderByDescending(x => x.Count)
                .Take(8)
                .ToList();

            var typeLines = new List<(string text, int fontSize, bool bold, string color)>
            {
                ("Distribución por Tipo de Siniestro", 22, true, "0C2340")
            };
            foreach (var t in typeTop)
            {
                var pct = totalAvisos > 0 ? (double)t.Count / totalAvisos * 100 : 0;
                var nombre = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(t.Tipo?.ToLower() ?? "");
                typeLines.Add(($"• {nombre}: {t.Count} ({pct:F1}%)", 12, false, "333333"));
            }
            CreateTextSlide(presentationPart, ref id, typeLines.ToArray());

            // --- Slide 5: Summary ---
            var topDeptText = deptTop.Count > 0
                ? $"• El departamento con más avisos es {CultureInfo.CurrentCulture.TextInfo.ToTitleCase(deptTop[0].Dept.ToLower())} ({deptTop[0].Count})."
                : "";
            var summaryLines = new List<(string text, int fontSize, bool bold, string color)>
            {
                ("Resumen y Conclusiones", 22, true, "0C2340"),
                ($"• Se registraron {totalAvisos.ToString("N0", EsPe)} avisos de siniestro en el ámbito analizado.", 12, false, "333333"),
                ($"• El monto total de indemnizaciones asciende a S/ {montoIndemn.ToString("N2", EsPe)}.", 12, false, "333333"),
                ($"• Se han desembolsado S/ {montoDesemb.ToString("N2", EsPe)} ({pctDesemb:F1}%) a {productores.ToString("N0", EsPe)} productores.", 12, false, "333333"),
            };
            if (!string.IsNullOrEmpty(topDeptText))
                summaryLines.Add((topDeptText, 12, false, "333333"));
            summaryLines.Add(("MIDAGRI — Dirección General de Seguimiento y Evaluación de Políticas", 9, false, "999999"));
            CreateTextSlide(presentationPart, ref id, summaryLines.ToArray());

            presentationPart.Presentation.Save();
        }

        return ms.ToArray();
    }

    public byte[] GeneratePptHistorico(DatosNacionalesDto datos, string departamento)
    {
        var deptUpper = departamento.Trim().ToUpper();
        var items = datos.Midagri.Where(s => s.Departamento == deptUpper).ToList();
        var deptTitle = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(deptUpper.ToLower());

        var mat = datos.Materia.FirstOrDefault(m =>
            m.Departamento.Equals(deptUpper, StringComparison.OrdinalIgnoreCase));
        var primaNeta = mat?.PrimaNeta ?? 0;
        var empresa = mat?.EmpresaAseguradora ?? "N/D";

        int totalAvisos = items.Count;
        double montoIndemn = items.Sum(s => s.Indemnizacion);
        double haIndemn = items.Sum(s => s.SupIndemnizada);
        double montoDesemb = items.Sum(s => s.MontoDesembolsado);
        double siniestralidad = primaNeta > 0 ? montoIndemn / primaNeta * 100 : 0;

        using var ms = new MemoryStream();
        using (var ppt = PresentationDocument.Create(ms, PresentationDocumentType.Presentation, true))
        {
            var presentationPart = ppt.AddPresentationPart();
            presentationPart.Presentation = new Presentation();
            presentationPart.Presentation.AppendChild(new SlideIdList());
            presentationPart.Presentation.AppendChild(
                new SlideSize { Cx = 12192000, Cy = 6858000 });
            presentationPart.Presentation.AppendChild(
                new NotesSize { Cx = 6858000, Cy = 9144000 });

            uint id = 256;

            // Slide 1: Title
            CreateTextSlide(presentationPart, ref id, new[]
            {
                ("ANÁLISIS HISTÓRICO DE SINIESTRALIDAD", 24, true, "0C2340"),
                ($"Departamento de {deptTitle}", 18, true, "408B14"),
                ("SAC 2025-2026 — Campaña Actual", 14, false, "333333"),
                ($"Aseguradora: {empresa}", 12, false, "7F8C8D"),
                ($"Corte: {datos.FechaCorte}", 11, false, "999999"),
                ("MIDAGRI", 10, false, "999999"),
            });

            // Slide 2: Current campaign KPIs
            var sinColor = siniestralidad > 70 ? "E74C3C"
                         : siniestralidad > 50 ? "F39C12"
                         : "27AE60";
            CreateTextSlide(presentationPart, ref id, new[]
            {
                ($"Campaña 2025-2026 — {deptTitle}", 22, true, "0C2340"),
                ($"Total Avisos: {totalAvisos.ToString("N0", EsPe)}", 14, true, "2980B9"),
                ($"Hectáreas Indemnizadas: {haIndemn.ToString("N2", EsPe)}", 13, false, "333333"),
                ($"Indemnización: S/ {montoIndemn.ToString("N2", EsPe)}", 14, true, "E67E22"),
                ($"Prima Neta: S/ {primaNeta.ToString("N2", EsPe)}", 13, false, "333333"),
                ($"Índice de Siniestralidad: {siniestralidad:F1}%", 16, true, sinColor),
                ($"Desembolsos: S/ {montoDesemb.ToString("N2", EsPe)}", 13, false, "333333"),
            });

            // Slide 3: Disaster type breakdown
            var tipos = items
                .GroupBy(s => s.TipoSiniestro)
                .Select(g => (Tipo: g.Key, Count: g.Count()))
                .OrderByDescending(x => x.Count)
                .Take(6)
                .ToList();

            var tipoLines = new List<(string text, int fontSize, bool bold, string color)>
            {
                ($"Tipos de Siniestro — {deptTitle}", 22, true, "0C2340")
            };
            foreach (var t in tipos)
            {
                var pct = totalAvisos > 0 ? (double)t.Count / totalAvisos * 100 : 0;
                var nombre = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(t.Tipo?.ToLower() ?? "");
                tipoLines.Add(($"• {nombre}: {t.Count} avisos ({pct:F1}%)", 12, false, "333333"));
            }
            CreateTextSlide(presentationPart, ref id, tipoLines.ToArray());

            // Slide 4: Historical context note
            CreateTextSlide(presentationPart, ref id, new[]
            {
                ("Contexto Histórico", 22, true, "0C2340"),
                ("Nota: Los datos históricos de campañas anteriores (2020-2025)", 12, false, "333333"),
                ("se encuentran disponibles en los archivos estáticos de la aplicación.", 12, false, "333333"),
                ("La integración completa con gráficos de evolución multi-campaña", 12, false, "333333"),
                ("se habilitará en una actualización futura.", 12, false, "333333"),
                ("MIDAGRI — DGASFS / DSFFA", 9, false, "999999"),
            });

            presentationPart.Presentation.Save();
        }

        return ms.ToArray();
    }

    // ================================================================
    //  Private helpers
    // ================================================================

    private static List<Siniestro> ApplyFilters(List<Siniestro> data, PptFilterDto? f)
    {
        if (f == null) return data;

        var result = data.AsEnumerable();

        if (!string.IsNullOrEmpty(f.Empresa) && f.Empresa != "Ambas")
        {
            // Placeholder: empresa filtering requires cross-reference with MateriaAsegurada
            // which maps departments to insurers. For now, pass-through.
        }

        if (f.Departamentos.Count > 0)
            result = result.Where(s =>
                f.Departamentos.Any(d => d.Equals(s.Departamento, StringComparison.OrdinalIgnoreCase)));

        if (f.Provincias.Count > 0)
            result = result.Where(s =>
                f.Provincias.Any(p => p.Equals(s.Provincia, StringComparison.OrdinalIgnoreCase)));

        if (f.Distritos.Count > 0)
            result = result.Where(s =>
                f.Distritos.Any(d => d.Equals(s.Distrito, StringComparison.OrdinalIgnoreCase)));

        if (f.TiposSiniestro.Count > 0)
            result = result.Where(s =>
                f.TiposSiniestro.Any(t => t.Equals(s.TipoSiniestro, StringComparison.OrdinalIgnoreCase)));

        if (f.FiltrarFecha && f.FechaInicio.HasValue && f.FechaFin.HasValue)
            result = result.Where(s =>
                s.FechaAviso >= f.FechaInicio && s.FechaAviso <= f.FechaFin);

        return result.ToList();
    }

    private static string DetermineScope(PptFilterDto? f)
    {
        if (f == null) return "Nacional";
        if (f.Distritos.Count > 0) return "Distrital";
        if (f.Provincias.Count > 0) return "Provincial";
        if (f.Departamentos.Count > 0) return "Departamental";
        return "Nacional";
    }

    private static string BuildScopeDetail(PptFilterDto? f)
    {
        if (f == null) return "Todos los departamentos";

        if (f.Distritos.Count > 0)
            return string.Join(", ", f.Distritos.Select(d =>
                CultureInfo.CurrentCulture.TextInfo.ToTitleCase(d.ToLower())));

        if (f.Provincias.Count > 0)
            return string.Join(", ", f.Provincias.Select(p =>
                CultureInfo.CurrentCulture.TextInfo.ToTitleCase(p.ToLower())));

        if (f.Departamentos.Count > 0)
            return string.Join(", ", f.Departamentos.Select(d =>
                CultureInfo.CurrentCulture.TextInfo.ToTitleCase(d.ToLower())));

        return "Todos los departamentos";
    }

    /// <summary>
    /// Creates a slide with a single text-box containing multiple paragraphs.
    /// Each line tuple: (text, fontSizePt, bold, hexColor).
    /// </summary>
    private static void CreateTextSlide(
        PresentationPart presPart,
        ref uint slideId,
        (string text, int fontSize, bool bold, string color)[] lines)
    {
        var slidePart = presPart.AddNewPart<SlidePart>();

        // Build the shape tree
        var shapeTree = new ShapeTree();

        // Required non-visual group shape properties
        shapeTree.Append(new NonVisualGroupShapeProperties(
            new NonVisualDrawingProperties { Id = 1U, Name = "" },
            new NonVisualGroupShapeDrawingProperties(),
            new ApplicationNonVisualDrawingProperties()));
        shapeTree.Append(new GroupShapeProperties(new D.TransformGroup()));

        // Content text box
        var shape = new Shape();

        shape.Append(new NonVisualShapeProperties(
            new NonVisualDrawingProperties { Id = 2U, Name = "Content" },
            new NonVisualShapeDrawingProperties(
                new D.ShapeLocks { NoGrouping = true }),
            new ApplicationNonVisualDrawingProperties()));

        shape.Append(new ShapeProperties(
            new D.Transform2D(
                new D.Offset { X = 457200L, Y = 457200L },
                new D.Extents { Cx = 11277600L, Cy = 5943600L }),
            new D.PresetGeometry(new D.AdjustValueList())
            { Preset = D.ShapeTypeValues.Rectangle }));

        var textBody = new TextBody(
            new D.BodyProperties
            {
                Wrap = D.TextWrappingValues.Square,
                Anchor = D.TextAnchoringTypeValues.Top
            },
            new D.ListStyle());

        foreach (var (text, fontSize, isBold, color) in lines)
        {
            if (string.IsNullOrEmpty(text)) continue;

            var paragraph = new D.Paragraph();

            // Paragraph properties with spacing
            var paraProps = new D.ParagraphProperties();
            paraProps.Append(new D.SpaceAfter(
                new D.SpacingPoints { Val = 600 }));
            paragraph.Append(paraProps);

            // Run with styled text
            var runProps = new D.RunProperties
            {
                Language = "es-PE",
                FontSize = fontSize * 100,
                Bold = isBold
            };
            runProps.Append(new D.SolidFill(
                new D.RgbColorModelHex { Val = color }));
            runProps.Append(new D.LatinFont { Typeface = "Calibri" });

            var run = new D.Run();
            run.Append(runProps);
            run.Append(new D.Text(text));

            paragraph.Append(run);
            textBody.Append(paragraph);
        }

        shape.Append(textBody);
        shapeTree.Append(shape);

        slidePart.Slide = new Slide(new CommonSlideData(shapeTree));

        // Register slide in the presentation
        var slideIdList = presPart.Presentation.SlideIdList!;
        slideIdList.Append(new SlideId
        {
            Id = slideId++,
            RelationshipId = presPart.GetIdOfPart(slidePart)
        });
    }
}
