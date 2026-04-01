namespace Infrastructure.DocumentGeneration;

using Application.Contracts.Services;
using Application.DTOs.Responses;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

public class WordDocumentGenerator : IWordReportService
{
    private static readonly CultureInfo EsPe = new("es-PE");

    private static string FmtNum(double val, int dec = 2)
    {
        if (dec == 0) return val.ToString("N0", EsPe);
        return val.ToString($"N{dec}", EsPe);
    }

    public byte[] GenerateNacionalDocx(DatosNacionalesDto datos)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            // Page margins
            var sectionProps = new SectionProperties(
                new PageMargin { Top = 1200, Right = 1200, Bottom = 1200, Left = 1200 }
            );

            // === TITLE ===
            AddCenteredParagraph(body, "AYUDA MEMORIA: RESUMEN OPERATIVIDAD SAC 2025-2026",
                14, true, "2F5496");
            AddCenteredParagraph(body, $"(al {datos.FechaCorte})", 10, false, "666666");
            AddEmptyParagraph(body);

            // === ACTIVACIÓN DEL SAC ===
            AddHeading(body, "Activación del SAC", "2F5496");
            AddParagraph(body, "El Seguro Agrícola Catastrófico se activa mediante el siguiente procedimiento:", 11);

            var pasos = new[]
            {
                "El productor reporta el siniestro a la Agencia/Oficina Agraria (presencial o telefónicamente).",
                "La Agencia Agraria, vía la DRA, notifica a la aseguradora por correo electrónico dentro de 7 días calendario.",
                "La aseguradora designa un perito ajustador para evaluar daños en campo dentro de 15 días calendario.",
                "El perito evalúa los daños con un agente agrario y elabora el acta de ajuste.",
                "Si se confirma pérdida catastrófica, se coordina el empadronamiento de agricultores dentro de 20 días calendario.",
                "La aseguradora abre cuentas bancarias y paga S/ 1,000 por hectárea asegurada dentro de 15 días hábiles tras la aprobación del padrón."
            };
            for (int i = 0; i < pasos.Length; i++)
                AddNumberedItem(body, i + 1, pasos[i]);
            AddEmptyParagraph(body);

            // === DATOS GENERALES ===
            var m = datos.Metricas;
            AddHeading(body, "Datos Generales a Nivel Nacional", "2F5496");

            var bullets = new[]
            {
                $"Empresas aseguradoras: {datos.EmpresasText}.",
                $"Prima total (con IGV): S/ {FmtNum(m.PrimaTotal)} | Prima neta (sin IGV): S/ {FmtNum(m.PrimaNeta)}",
                $"Superficie asegurada: {FmtNum(m.SupAsegurada)} hectáreas en 24 departamentos.",
                $"Productores asegurados (estimados): {FmtNum(m.ProdAsegurados, 0)} / Suma asegurada por hectárea: S/ 1.000,00",
                $"Avisos de siniestros: {FmtNum(m.TotalAvisos, 0)} reportados | {FmtNum(m.TotalAjustados, 0)} ajustados ({m.PctAjustados:F1}%)",
                $"Indemnizaciones reconocidas: S/ {FmtNum(m.MontoIndemnizado)} | Índice de siniestralidad: {m.IndiceSiniestralidad:F2}%",
                $"Desembolsos realizados: S/ {FmtNum(m.MontoDesembolsado)} ({m.PctDesembolso:F1}%) a {FmtNum(m.ProductoresDesembolso, 0)} productores en {m.DeptosConDesembolso} de 24 departamentos."
            };
            foreach (var b in bullets) AddBulletItem(body, b);
            AddEmptyParagraph(body);

            // === CUADRO 1 ===
            AddHeading(body, "Cuadro 1: Primas y Cobertura por Departamento", "2F5496");
            if (datos.Cuadro1.Count > 0)
            {
                var headers1 = new[] { "Departamento", "Prima Total (S/)", "Hectáreas Aseguradas", "Suma Asegurada Máxima (S/)" };
                var rows1 = datos.Cuadro1.Select(r => new[] { r.Departamento, FmtNum(r.PrimaTotal), FmtNum(r.Hectareas), FmtNum(r.SumaAsegurada) }).ToList();
                // Add TOTAL row
                rows1.Add(new[] { "TOTAL", FmtNum(datos.Cuadro1.Sum(r => r.PrimaTotal)), FmtNum(datos.Cuadro1.Sum(r => r.Hectareas)), FmtNum(datos.Cuadro1.Sum(r => r.SumaAsegurada)) });
                AddTable(body, headers1, rows1, "2F5496");
            }
            AddEmptyParagraph(body);

            // === CUADRO 2 ===
            AddHeading(body, "Cuadro 2: Indemnizaciones y Desembolsos por Departamento", "2F5496");
            if (datos.Cuadro2.Count > 0)
            {
                var headers2 = new[] { "Departamento", "Ha Indemnizadas", "Monto Indemnizado (S/)", "Monto Desembolsado (S/)", "Productores con Desembolso" };
                var rows2 = datos.Cuadro2.Select(r => new[] { r.Departamento, FmtNum(r.HaIndemnizadas), FmtNum(r.MontoIndemnizado), FmtNum(r.MontoDesembolsado), FmtNum(r.Productores, 0) }).ToList();
                rows2.Add(new[] { "TOTAL", FmtNum(datos.Cuadro2.Sum(r => r.HaIndemnizadas)), FmtNum(datos.Cuadro2.Sum(r => r.MontoIndemnizado)), FmtNum(datos.Cuadro2.Sum(r => r.MontoDesembolsado)), FmtNum(datos.Cuadro2.Sum(r => r.Productores), 0) });
                AddTable(body, headers2, rows2, "2F5496");
            }
            AddEmptyParagraph(body);

            // === CUADRO 3 ===
            AddHeading(body, "Cuadro 3: Eventos Asociados a Lluvias Intensas por Departamento", "2F5496");

            // Descriptive text about lluvia
            var lluviaDesc = string.Join(", ", datos.LluviaPorTipo.Select(kvp => $"{kvp.Key.ToLower()} ({kvp.Value})"));
            var top3Lluvia = datos.Cuadro3.Where(r => r.Departamento != "TOTAL").OrderByDescending(r => r.Avisos).Take(3)
                .Select(r => $"{CultureInfo.CurrentCulture.TextInfo.ToTitleCase(r.Departamento.ToLower())} ({r.Avisos} avisos)");
            AddParagraph(body, $"Se registran {FmtNum(datos.TotalLluvia, 0)} avisos por eventos asociados a lluvias intensas ({datos.PctLluvia:F1}% del total), que incluyen {lluviaDesc}. Los departamentos más afectados son {string.Join(", ", top3Lluvia)}.", 11);

            if (datos.Cuadro3.Count > 0)
            {
                var headers3 = new[] { "Departamento", "Avisos", "Ha Indemn.", "Monto Indemnizado (S/)", "Monto Desembolsado (S/)", "Productores" };
                var rows3 = datos.Cuadro3.Select(r => new[] { r.Departamento, r.Avisos.ToString(), FmtNum(r.HaIndemn), FmtNum(r.MontoIndemnizado), FmtNum(r.MontoDesembolsado), FmtNum(r.Productores, 0) }).ToList();
                AddTable(body, headers3, rows3, "2F5496");
            }
            AddEmptyParagraph(body);

            // === NOTA FINAL ===
            var notePara = new Paragraph();
            var noteRun1 = CreateRun("Nota: ", 9, true, "000000", true);
            var noteRun2 = CreateRun("La vigencia de la póliza es del 01/08/2025 al 01/08/2026. ", 9, false, "000000", true);
            // Top 3 siniestros text
            string top3Text = "";
            if (datos.Top3Siniestros.Count > 0)
            {
                var parts = datos.Top3Siniestros.Select(kvp => {
                    var pct = datos.Metricas.TotalAvisos > 0 ? (double)kvp.Value / datos.Metricas.TotalAvisos * 100 : 0;
                    return $"{kvp.Key.ToLower()} ({pct:F1}%)";
                });
                top3Text = $"Los siniestros principales son {string.Join(", ", parts)}.";
            }
            var noteRun3 = CreateRun(top3Text, 9, false, "000000", true);
            notePara.Append(noteRun1, noteRun2, noteRun3);
            body.Append(notePara);

            body.Append(sectionProps);
        }

        return ms.ToArray();
    }

    public byte[] GenerateDepartamentalDocx(DatosDepartamentalDto d)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            var sectionProps = new SectionProperties(
                new PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 }
            );

            // === COVER PAGE ===
            for (int i = 0; i < 3; i++) AddEmptyParagraph(body);
            AddCenteredParagraph(body, "AYUDA MEMORIA", 28, true, "2E4057");
            AddCenteredParagraph(body, "Seguro Agrícola Catastrófico (SAC)", 15, false, "C0392B");
            AddEmptyParagraph(body);
            AddCenteredParagraph(body, $"Departamento de {d.Departamento}", 14, true, "000000");
            AddCenteredParagraph(body, "Campaña 2025 - 2026", 11, false, "000000");
            AddCenteredParagraph(body, $"Prima Neta Departamental: S/ {FmtNum(d.PrimaNeta)}", 11, false, "000000");
            AddCenteredParagraph(body, $"Corte de datos: {d.FechaCorte}", 11, false, "7F8C8D");
            AddEmptyParagraph(body);
            AddEmptyParagraph(body);
            AddCenteredParagraph(body, "MIDAGRI", 10, false, "7F8C8D");

            // Page break
            body.Append(new Paragraph(new Run(new Break { Type = BreakValues.Page })));

            // === INTRODUCCIÓN ===
            AddHeading(body, "Introducción", "2E4057");
            AddParagraph(body, $"La aseguradora responsable del departamento es {d.Empresa}.", 11);
            AddParagraph(body, $"Para la campaña 2025-2026, {d.Departamento} cuenta con una cobertura de {FmtNum(d.SupAsegurada, 0)} hectáreas aseguradas, por un monto de prima neta de S/ {FmtNum(d.PrimaNeta)}. El SAC 2025-2026 ha incrementado la indemnización máxima de S/ 800 a S/ 1,000 por hectárea de cultivo afectado.", 11);
            AddEmptyParagraph(body);

            // === PROCESO SAC ===
            AddHeading(body, "1. Proceso del SAC: Etapas y Plazos", "2E4057");
            AddParagraph(body, "El Seguro Agrícola Catastrófico (SAC) sigue un proceso de 8 etapas desde la ocurrencia del siniestro hasta el desembolso al productor afectado.", 10);

            var etapas = new[]
            {
                new[] { "1", "Ocurrencia del Siniestro", "Evento climático adverso afecta cultivos", "—" },
                new[] { "2", "Aviso de Siniestro", "El productor o DRAS comunica el evento", "Productor / DRAS" },
                new[] { "3", "Atención del Aviso", "Aseguradora registra y verifica", "Aseguradora" },
                new[] { "4", "Programación de Ajuste", "Se coordina visita de inspección", "Aseguradora" },
                new[] { "5", "Inspección y Ajuste", "Evaluación técnica en campo", "Ajustador" },
                new[] { "6", "Dictamen", "Se determina si es indemnizable", "Aseguradora" },
                new[] { "7", "Validación DRAS/GRAS", "DRA emite conformidad", "DRAS/GRAS" },
                new[] { "8", "Desembolso", "Pago de indemnización al productor", "Aseguradora" },
            };
            AddTable(body, new[] { "N°", "Etapa", "Descripción", "Responsable" }, etapas.ToList(), "2C3E50");
            AddEmptyParagraph(body);

            // === PANORAMA GENERAL ===
            AddHeading(body, $"2. Panorama General — {d.Departamento}", "2E4057");

            // Metrics as simple paragraphs (2x2 dashboard equivalent)
            AddParagraph(body, $"• Avisos de Siniestro registrados: {FmtNum(d.TotalAvisos, 0)}", 11);
            AddParagraph(body, $"• Indemnización total ({d.Indemnizables} casos indemnizables): S/ {FmtNum(d.MontoIndemnizado, 0)}", 11);
            AddParagraph(body, $"• Superficie indemnizada: {FmtNum(d.HaIndemnizadas)} ha (de {FmtNum(d.SupAsegurada, 0)} ha aseguradas)", 11);
            if (d.MontoDesembolsado > 0)
                AddParagraph(body, $"• Monto desembolsado a {d.ProductoresDesembolso} productores: S/ {FmtNum(d.MontoDesembolsado, 0)}", 11);
            else
                AddParagraph(body, "• Monto desembolsado (pendiente de pago)", 11);
            AddEmptyParagraph(body);

            // Avisos por Tipo
            AddSubHeading(body, "Avisos por Tipo de Siniestro", "16A085");
            if (d.AvisosTipo.Count > 0)
                AddTable(body, new[] { "Tipo de Siniestro", "N° Avisos", "% del Total" }, d.AvisosTipo.Select(a => a).ToList(), "2C3E50");
            AddEmptyParagraph(body);

            // Distribución por Provincia
            AddSubHeading(body, "Distribución por Provincia", "16A085");
            if (d.DistProvincia.Count > 0)
                AddTable(body, d.DistProvinciaHeaders.ToArray(), d.DistProvincia.Select(a => a).ToList(), "2C3E50");
            AddEmptyParagraph(body);

            // === EVENTOS RECIENTES ===
            AddHeading(body, "3. Eventos Registrados Recientemente", "2E4057");
            if (d.EventosRecientes.Count > 0)
            {
                AddParagraph(body, $"Se han registrado {d.EventosRecientes.Count} avisos de siniestro recientes en el departamento de {d.Departamento}.", 10);
                AddTable(body, d.EventosHeaders.ToArray(), d.EventosRecientes.Select(a => a).ToList(), "2C3E50");
            }
            else
            {
                AddParagraph(body, "No se han registrado eventos recientes en el período.", 10);
            }
            AddEmptyParagraph(body);

            // Disclaimer
            AddParagraph(body, "* Superficie perdida reportada preliminarmente; la superficie afectada total está pendiente de evaluación en campo.", 9, "7F8C8D");
            AddEmptyParagraph(body);

            // === RESUMEN OPERATIVO ===
            AddSubHeading(body, "Resumen Operativo", "16A085");
            if (!string.IsNullOrEmpty(d.ResumenOperativo))
                AddParagraph(body, d.ResumenOperativo, 10);
            if (!string.IsNullOrEmpty(d.ResumenDesembolso))
                AddParagraph(body, d.ResumenDesembolso, 10);

            body.Append(sectionProps);
        }

        return ms.ToArray();
    }

    // === HELPER METHODS ===

    private static Run CreateRun(string text, int fontSize, bool bold, string color, bool italic = false)
    {
        var run = new Run();
        var props = new RunProperties();
        props.Append(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" });
        props.Append(new FontSize { Val = (fontSize * 2).ToString() });
        if (bold) props.Append(new Bold());
        if (italic) props.Append(new Italic());
        props.Append(new Color { Val = color });
        run.Append(props);
        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    private static void AddParagraph(Body body, string text, int fontSize, string color = "000000")
    {
        var para = new Paragraph();
        para.Append(CreateRun(text, fontSize, false, color));
        body.Append(para);
    }

    private static void AddCenteredParagraph(Body body, string text, int fontSize, bool bold, string color)
    {
        var para = new Paragraph();
        var pProps = new ParagraphProperties(new Justification { Val = JustificationValues.Center });
        para.Append(pProps);
        para.Append(CreateRun(text, fontSize, bold, color));
        body.Append(para);
    }

    private static void AddHeading(Body body, string text, string color)
    {
        var para = new Paragraph();
        var pProps = new ParagraphProperties(new SpacingBetweenLines { Before = "240", After = "120" });
        para.Append(pProps);
        para.Append(CreateRun(text, 14, true, color));
        body.Append(para);
    }

    private static void AddSubHeading(Body body, string text, string color)
    {
        var para = new Paragraph();
        var pProps = new ParagraphProperties(new SpacingBetweenLines { Before = "200", After = "100" });
        para.Append(pProps);
        para.Append(CreateRun(text, 13, true, color));
        body.Append(para);
    }

    private static void AddEmptyParagraph(Body body)
    {
        body.Append(new Paragraph());
    }

    private static void AddNumberedItem(Body body, int number, string text)
    {
        var para = new Paragraph();
        para.Append(CreateRun($"{number}. {text}", 11, true, "000000"));
        body.Append(para);
    }

    private static void AddBulletItem(Body body, string text)
    {
        var para = new Paragraph();
        para.Append(CreateRun($"• {text}", 11, true, "000000"));
        body.Append(para);
    }

    private static void AddTable(Body body, string[] headers, List<string[]> rows, string headerBgColor)
    {
        var table = new Table();

        // Table properties
        var tblProps = new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4, Color = "BBBBBB" },
                new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "BBBBBB" },
                new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "BBBBBB" },
                new RightBorder { Val = BorderValues.Single, Size = 4, Color = "BBBBBB" },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "BBBBBB" },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "BBBBBB" }
            ),
            new TableWidth { Type = TableWidthUnitValues.Pct, Width = "5000" }
        );
        table.Append(tblProps);

        // Header row
        var headerRow = new TableRow();
        foreach (var h in headers)
        {
            var cell = new TableCell();
            var cellProps = new TableCellProperties(
                new Shading { Val = ShadingPatternValues.Clear, Fill = headerBgColor }
            );
            cell.Append(cellProps);
            var para = new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }));
            para.Append(CreateRun(h, 9, true, "FFFFFF"));
            cell.Append(para);
            headerRow.Append(cell);
        }
        table.Append(headerRow);

        // Data rows
        for (int i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            var isTotal = row.Length > 0 && row[0].ToUpper() == "TOTAL";
            var isAlt = i % 2 == 0;
            var dataRow = new TableRow();

            for (int j = 0; j < row.Length; j++)
            {
                var cell = new TableCell();
                var cellProps = new TableCellProperties();

                if (isTotal)
                    cellProps.Append(new Shading { Val = ShadingPatternValues.Clear, Fill = headerBgColor });
                else if (isAlt)
                    cellProps.Append(new Shading { Val = ShadingPatternValues.Clear, Fill = "D6E4F0" });

                cell.Append(cellProps);

                var align = j > 0 ? JustificationValues.Right : JustificationValues.Left;
                var para = new Paragraph(new ParagraphProperties(new Justification { Val = align }));
                var textColor = isTotal ? "FFFFFF" : "000000";
                para.Append(CreateRun(row[j] ?? "", 9, isTotal, textColor));
                cell.Append(para);
                dataRow.Append(cell);
            }
            table.Append(dataRow);
        }

        body.Append(table);
    }
}
