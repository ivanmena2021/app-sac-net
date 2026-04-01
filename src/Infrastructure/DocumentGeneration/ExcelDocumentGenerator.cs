namespace Infrastructure.DocumentGeneration;

using Application.Contracts.Services;
using Application.DTOs.Responses;
using ClosedXML.Excel;
using System.Globalization;

public class ExcelDocumentGenerator : IExcelReportService
{
    public byte[] GenerateReporteEme(DatosNacionalesDto datos)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Hoja1");

        // Headers
        var headers = new[]
        {
            "REGIÓN", "DISTRITOS", "BENEFICIARIOS",
            "ACCIÓN IMPLEMENTADA\n- MONTO DESEMBOLSADO",
            "ACCION EN IMPLEMENTACIÓN\n - MONTO INDEMNIZADO ",
            "ACCION POR IMPLEMENTAR - PRIMA TOTAL",
            "DESCRIPCIÓN", "UNIDAD RESPONSABLE",
            "CUANTIFICACIÓN/ TOTAL - HAS INDEMNIZADAS",
            "OBSERVACIONES"
        };

        // Header style
        for (int i = 0; i < headers.Length; i++)
        {
            var cell = ws.Cell(1, i + 1);
            cell.Value = headers[i];
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontSize = 10;
            cell.Style.Font.FontColor = XLColor.White;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1F4E79");
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            cell.Style.Alignment.WrapText = true;
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        // Group data by department
        var deptGroups = datos.Midagri
            .GroupBy(s => s.Departamento)
            .OrderBy(g => g.Key)
            .ToList();

        // Build materia lookup
        var materiaLookup = datos.Materia.ToDictionary(m => m.Departamento, m => m, StringComparer.OrdinalIgnoreCase);

        int row = 2;
        foreach (var group in deptGroups)
        {
            var depto = group.Key;
            var items = group.ToList();

            // District info
            var distInfo = items.GroupBy(s => s.Provincia)
                .OrderByDescending(g => g.Count())
                .Select(g => $"{ToTitleCase(g.Key)} ({g.Select(s => s.Distrito).Distinct().Count()})")
                .ToList();
            var nDist = items.Select(s => s.Distrito).Distinct().Count();
            var nProv = items.Select(s => s.Provincia).Distinct().Count();
            var distText = $"{nDist} distritos en {nProv} provincias: {string.Join(", ", distInfo)}";

            // Aggregates
            var nProductores = items.Sum(s => s.NProductores);
            var montoDesembolsado = items.Sum(s => s.MontoDesembolsado);
            var indemnizacion = items.Sum(s => s.Indemnizacion);
            var supIndemnizada = items.Sum(s => s.SupIndemnizada);

            // Static data
            materiaLookup.TryGetValue(depto, out var mat);
            var empresa = mat?.EmpresaAseguradora ?? "N/D";
            var primaTotal = mat?.PrimaTotal ?? 0;

            // Top 3 siniestros
            var topSin = items.GroupBy(s => s.TipoSiniestro)
                .OrderByDescending(g => g.Count())
                .Take(3)
                .Select(g => $"{g.Key.ToLower()} ({g.Count()})")
                .ToList();
            var tiposText = topSin.Count > 0 ? string.Join(", ", topSin) : "sin datos";
            var descripcion = $"Aseguradora: {empresa}. {items.Count} avisos de siniestro. Principales siniestros: {tiposText}.";

            // Observations
            string obs;
            if (montoDesembolsado > 0)
                obs = $"Desembolsos realizados a {(int)nProductores} productores.";
            else if (indemnizacion > 0)
                obs = "Indemnizaciones reconocidas, pendiente de desembolso.";
            else
                obs = "En proceso de evaluación.";

            // Write row
            ws.Cell(row, 1).Value = ToTitleCase(depto);
            ws.Cell(row, 2).Value = distText;
            ws.Cell(row, 3).Value = nProductores > 0 ? (int)nProductores : 0;
            ws.Cell(row, 4).Value = Math.Round(montoDesembolsado, 2);
            ws.Cell(row, 5).Value = Math.Round(indemnizacion, 2);
            ws.Cell(row, 6).Value = Math.Round(primaTotal, 2);
            ws.Cell(row, 7).Value = descripcion;
            ws.Cell(row, 8).Value = "DGASFS / DSFFA";
            ws.Cell(row, 9).Value = Math.Round(supIndemnizada, 2);
            ws.Cell(row, 10).Value = obs;

            // Style data row
            for (int c = 1; c <= 10; c++)
            {
                var cell = ws.Cell(row, c);
                cell.Style.Font.FontSize = 9;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                cell.Style.Alignment.WrapText = true;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
            }

            // Number format
            ws.Cell(row, 3).Style.NumberFormat.Format = "#,##0";
            ws.Cell(row, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            foreach (var c in new[] { 4, 5, 6, 9 })
            {
                ws.Cell(row, c).Style.NumberFormat.Format = "#,##0.00";
                ws.Cell(row, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            }

            row++;
        }

        // Column widths
        var colWidths = new[] { 15, 50, 13, 18, 18, 18, 50, 18, 18, 40 };
        for (int i = 0; i < colWidths.Length; i++)
            ws.Column(i + 1).Width = colWidths[i];

        ws.Row(1).Height = 40;

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        return ms.ToArray();
    }

    private static string ToTitleCase(string s)
    {
        if (string.IsNullOrEmpty(s)) return s;
        return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(s.ToLower());
    }
}
