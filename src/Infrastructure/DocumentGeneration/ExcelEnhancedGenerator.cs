namespace Infrastructure.DocumentGeneration;

using Application.Contracts.Services;
using Application.DTOs.Responses;
using ClosedXML.Excel;
using System.Globalization;

public class ExcelEnhancedGenerator : IExcelEnhancedService
{
    private static readonly CultureInfo EsPe = new("es-PE");

    public byte[] GenerateEnhancedExcel(DatosNacionalesDto datos)
    {
        using var workbook = new XLWorkbook();
        var m = datos.Metricas;

        // === Sheet 1: Resumen ===
        var wsResumen = workbook.Worksheets.Add("Resumen");

        // Title
        wsResumen.Range("A1:H1").Merge().Value = "Seguro Agrícola Catastrófico — SAC 2025-2026";
        wsResumen.Cell("A1").Style.Font.FontSize = 14;
        wsResumen.Cell("A1").Style.Font.Bold = true;
        wsResumen.Cell("A1").Style.Font.FontColor = XLColor.FromHtml("#0C2340");
        wsResumen.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        wsResumen.Row(1).Height = 35;

        wsResumen.Range("A2:H2").Merge().Value = $"Fecha de corte: {datos.FechaCorte}";
        wsResumen.Cell("A2").Style.Font.FontSize = 10;
        wsResumen.Cell("A2").Style.Font.Italic = true;
        wsResumen.Cell("A2").Style.Font.FontColor = XLColor.FromHtml("#1A5276");
        wsResumen.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        // KPI Row 1 (labels in row 3, values in row 4)
        var kpis = new (string label, object value, string format)[]
        {
            ("Total Avisos", m.TotalAvisos, "#,##0"),
            ("Ha Indemnizadas", m.HaIndemnizadas, "#,##0.00"),
            ("Monto Indemnizado (S/)", m.MontoIndemnizado, "#,##0.00"),
            ("Siniestralidad (%)", m.IndiceSiniestralidad / 100.0, "0.0%"),
            ("Monto Desembolsado (S/)", m.MontoDesembolsado, "#,##0.00"),
            ("Productores", m.ProductoresDesembolso, "#,##0"),
            ("% Desembolso", m.PctDesembolso / 100.0, "0.0%"),
            ("Prima Total (S/)", m.PrimaTotal, "#,##0.00"),
        };

        for (int i = 0; i < kpis.Length; i++)
        {
            int col = (i % 4) * 2 + 1;
            int labelRow = i < 4 ? 3 : 5;
            int valueRow = labelRow + 1;

            var labelRange = wsResumen.Range(labelRow, col, labelRow, col + 1).Merge();
            labelRange.Value = kpis[i].label;
            labelRange.Style.Font.Bold = true;
            labelRange.Style.Font.FontSize = 10;
            labelRange.Style.Font.FontColor = XLColor.FromHtml("#1A5276");

            var valueRange = wsResumen.Range(valueRow, col, valueRow, col + 1).Merge();
            valueRange.FirstCell().SetValue(Convert.ToDouble(kpis[i].value));
            valueRange.Style.Font.Bold = true;
            valueRange.Style.Font.FontSize = 13;
            valueRange.Style.Font.FontColor = XLColor.FromHtml("#0C2340");
            valueRange.Style.NumberFormat.Format = kpis[i].format;
        }

        wsResumen.Row(3).Height = 22;
        wsResumen.Row(4).Height = 30;
        wsResumen.Row(5).Height = 22;
        wsResumen.Row(6).Height = 30;

        // Department ranking from Cuadro2
        int startRow = 8;
        wsResumen.Cell(startRow, 1).Value = "Departamento";
        wsResumen.Cell(startRow, 2).Value = "Ha Indemnizadas";
        wsResumen.Cell(startRow, 3).Value = "Monto Indemnizado (S/)";
        wsResumen.Cell(startRow, 4).Value = "Monto Desembolsado (S/)";
        wsResumen.Cell(startRow, 5).Value = "Productores";

        for (int c = 1; c <= 5; c++)
        {
            wsResumen.Cell(startRow, c).Style.Font.Bold = true;
            wsResumen.Cell(startRow, c).Style.Font.FontColor = XLColor.White;
            wsResumen.Cell(startRow, c).Style.Fill.BackgroundColor = XLColor.FromHtml("#0C2340");
            wsResumen.Cell(startRow, c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        var sortedCuadro = datos.Cuadro2.Where(r => r.Departamento != "TOTAL").OrderByDescending(r => r.MontoIndemnizado).ToList();
        for (int i = 0; i < sortedCuadro.Count; i++)
        {
            int row = startRow + 1 + i;
            var r = sortedCuadro[i];
            wsResumen.Cell(row, 1).Value = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(r.Departamento.ToLower());
            wsResumen.Cell(row, 2).Value = r.HaIndemnizadas;
            wsResumen.Cell(row, 3).Value = r.MontoIndemnizado;
            wsResumen.Cell(row, 4).Value = r.MontoDesembolsado;
            wsResumen.Cell(row, 5).Value = r.Productores;

            for (int c = 1; c <= 5; c++)
            {
                wsResumen.Cell(row, c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                if (i % 2 == 1)
                    wsResumen.Cell(row, c).Style.Fill.BackgroundColor = XLColor.FromHtml("#F8F9FA");
            }
            wsResumen.Cell(row, 2).Style.NumberFormat.Format = "#,##0.00";
            wsResumen.Cell(row, 3).Style.NumberFormat.Format = "#,##0.00";
            wsResumen.Cell(row, 4).Style.NumberFormat.Format = "#,##0.00";
            wsResumen.Cell(row, 5).Style.NumberFormat.Format = "#,##0";
        }

        wsResumen.Columns().AdjustToContents(10, 35);

        // === Sheet 2: Consolidado ===
        BuildDataSheet(workbook, datos.Midagri, "Consolidado");

        // === Sheet 3 & 4: By Company ===
        var lp = datos.Midagri.Where(s => s.TipoSiniestro != null).ToList(); // All data (we don't have EMPRESA field split yet)
        // For now, create a single "Todos" sheet since we don't track EMPRESA per siniestro in the entity
        // This can be enhanced later when EMPRESA field is added to Siniestro entity

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        return ms.ToArray();
    }

    private static void BuildDataSheet(XLWorkbook workbook, List<Domain.Entities.Siniestro> data, string sheetName)
    {
        var ws = workbook.Worksheets.Add(sheetName);
        var headers = new[] { "Departamento", "Provincia", "Distrito", "Tipo Cultivo", "Tipo Siniestro",
            "Estado Inspección", "Dictamen", "Sup. Indemnizada", "Indemnización", "Monto Desembolsado", "Productores" };

        for (int c = 0; c < headers.Length; c++)
        {
            ws.Cell(1, c + 1).Value = headers[c];
            ws.Cell(1, c + 1).Style.Font.Bold = true;
            ws.Cell(1, c + 1).Style.Font.FontColor = XLColor.White;
            ws.Cell(1, c + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#0C2340");
            ws.Cell(1, c + 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        for (int i = 0; i < Math.Min(data.Count, 10000); i++)
        {
            var s = data[i];
            int row = i + 2;
            ws.Cell(row, 1).Value = s.Departamento;
            ws.Cell(row, 2).Value = s.Provincia;
            ws.Cell(row, 3).Value = s.Distrito;
            ws.Cell(row, 4).Value = s.TipoCultivo;
            ws.Cell(row, 5).Value = s.TipoSiniestro;
            ws.Cell(row, 6).Value = s.EstadoInspeccion;
            ws.Cell(row, 7).Value = s.Dictamen;
            ws.Cell(row, 8).Value = s.SupIndemnizada;
            ws.Cell(row, 9).Value = s.Indemnizacion;
            ws.Cell(row, 10).Value = s.MontoDesembolsado;
            ws.Cell(row, 11).Value = s.NProductores;

            ws.Cell(row, 8).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(row, 9).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(row, 10).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(row, 11).Style.NumberFormat.Format = "#,##0";

            if (s.Indemnizacion > 0)
                for (int c = 1; c <= 11; c++)
                    ws.Cell(row, c).Style.Fill.BackgroundColor = XLColor.FromHtml("#D5F5E3");
            else if (i % 2 == 1)
                for (int c = 1; c <= 11; c++)
                    ws.Cell(row, c).Style.Fill.BackgroundColor = XLColor.FromHtml("#F8F9FA");
        }

        // Freeze header + auto-filter
        ws.SheetView.FreezeRows(1);
        ws.RangeUsed()?.SetAutoFilter();
        ws.Columns().AdjustToContents(10, 35);
    }
}
