namespace Infrastructure.Alertas;

using Application.Contracts.Services;
using Application.DTOs.Responses;

public class SemaforoService : ISemaforoService
{
    // Deadline thresholds in days for each stage
    private static readonly Dictionary<string, (int verde, int ambar)> Thresholds = new()
    {
        ["Atención"] = (7, 10),
        ["Programación"] = (10, 15),
        ["Ajuste"] = (15, 20),
        ["Validación"] = (15, 20),
        ["Pago"] = (15, 20),
    };

    public SemaforoResultDto ComputeSemaforo(DatosNacionalesDto datos)
    {
        var result = new SemaforoResultDto();
        var today = DateTime.Today;

        foreach (var s in datos.Midagri)
        {
            // Skip null avisos
            if (string.IsNullOrEmpty(s.CodigoAviso) || s.CodigoAviso.ToUpper().Contains("NULO"))
                continue;

            result.TotalAvisos++;

            // Determine current stage and calculate days
            string etapa;
            string alerta;
            int dias;
            string detalle;

            if (s.MontoDesembolsado > 0 && s.FechaDesembolso.HasValue)
            {
                // Already disbursed - complete
                etapa = "Completado";
                alerta = "verde";
                dias = 0;
                detalle = "Desembolso realizado";
            }
            else if (!string.IsNullOrEmpty(s.Dictamen) && s.Dictamen.ToUpper() != "PENDIENTE")
            {
                // Has dictamen but no disbursement yet
                etapa = "Pago";
                dias = s.FechaAviso.HasValue ? (int)(today - s.FechaAviso.Value).TotalDays : 0;
                alerta = ClassifyAlert("Pago", dias);
                detalle = $"Dictamen: {s.Dictamen}, pendiente de pago ({dias} días)";
            }
            else if (s.EstadoInspeccion?.ToUpper() == "CERRADO")
            {
                // Closed inspection, pending dictamen/payment
                etapa = "Validación";
                dias = s.FechaAviso.HasValue ? (int)(today - s.FechaAviso.Value).TotalDays : 0;
                alerta = ClassifyAlert("Validación", dias);
                detalle = $"Inspección cerrada, pendiente validación ({dias} días)";
            }
            else if (s.FechaAtencion.HasValue)
            {
                // Has been attended, pending inspection
                etapa = "Ajuste";
                dias = (int)(today - s.FechaAtencion.Value).TotalDays;
                alerta = ClassifyAlert("Ajuste", dias);
                detalle = $"En proceso de ajuste ({dias} días desde atención)";
            }
            else if (s.FechaAviso.HasValue)
            {
                // Has aviso date but no attention yet
                etapa = "Atención";
                dias = (int)(today - s.FechaAviso.Value).TotalDays;
                alerta = ClassifyAlert("Atención", dias);
                detalle = $"Pendiente de atención ({dias} días desde aviso)";
            }
            else
            {
                etapa = "Registro";
                dias = 0;
                alerta = "verde";
                detalle = "Sin fecha de aviso registrada";
            }

            // Count by alert level
            switch (alerta)
            {
                case "verde": result.Verde++; break;
                case "ambar": result.Ambar++; break;
                case "rojo": result.Rojo++; break;
            }

            result.Rows.Add(new SemaforoRow
            {
                CodigoAviso = s.CodigoAviso,
                Departamento = s.Departamento,
                Provincia = s.Provincia,
                Distrito = s.Distrito,
                Etapa = etapa,
                Alerta = alerta,
                Dias = dias,
                Detalle = detalle,
            });

            // Aggregate by department
            if (!result.PorDepartamento.ContainsKey(s.Departamento))
                result.PorDepartamento[s.Departamento] = new SemaforoDepartamento();

            var dept = result.PorDepartamento[s.Departamento];
            dept.Total++;
            switch (alerta)
            {
                case "verde": dept.Verde++; break;
                case "ambar": dept.Ambar++; break;
                case "rojo": dept.Rojo++; break;
            }
        }

        return result;
    }

    private static string ClassifyAlert(string etapa, int dias)
    {
        if (!Thresholds.TryGetValue(etapa, out var threshold))
            return "verde";

        if (dias <= threshold.verde) return "verde";
        if (dias <= threshold.ambar) return "ambar";
        return "rojo";
    }
}
