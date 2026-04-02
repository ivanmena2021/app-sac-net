namespace Application.Services;

using Application.Contracts.Repositories;
using Application.Contracts.Services;
using Application.DTOs.Responses;
using Domain.Entities;
using Domain.Enums;
using System.Globalization;
using System.Text;

public class DataProcessorService : IDataProcessorService
{
    private readonly IExcelReaderRepository _excelReader;
    private readonly IStaticDataRepository _staticData;

    private static readonly Dictionary<string, string> CapitalToDepto = new(StringComparer.OrdinalIgnoreCase)
    {
        { "CHACHAPOYAS", "AMAZONAS" },
        { "HUARAZ", "ANCASH" },
        { "ABANCAY", "APURIMAC" },
        { "AREQUIPA", "AREQUIPA" },
        { "AYACUCHO", "AYACUCHO" },
        { "CAJAMARCA", "CAJAMARCA" },
        { "CUSCO", "CUSCO" },
        { "HUANCAVELICA", "HUANCAVELICA" },
        { "HUANUCO", "HUANUCO" },
        { "ICA", "ICA" },
        { "HUANCAYO", "JUNIN" },
        { "TRUJILLO", "LA LIBERTAD" },
        { "CHICLAYO", "LAMBAYEQUE" },
        { "LIMA", "LIMA" },
        { "IQUITOS", "LORETO" },
        { "PUERTO MALDONADO", "MADRE DE DIOS" },
        { "MOQUEGUA", "MOQUEGUA" },
        { "CERRO DE PASCO", "PASCO" },
        { "PIURA", "PIURA" },
        { "PUNO", "PUNO" },
        { "MOYOBAMBA", "SAN MARTIN" },
        { "TACNA", "TACNA" },
        { "TUMBES", "TUMBES" },
        { "PUCALLPA", "UCAYALI" }
    };

    public DataProcessorService(IExcelReaderRepository excelReader, IStaticDataRepository staticData)
    {
        _excelReader = excelReader;
        _staticData = staticData;
    }

    // ────────────────────────────────────────────
    //  PUBLIC: ProcessDynamicData
    // ────────────────────────────────────────────
    public DatosNacionalesDto ProcessDynamicData(Stream midagriStream, Stream siniestrosStream)
    {
        // 1. Read Excel files
        List<Siniestro> midagri;
        try
        {
            midagri = _excelReader.ReadMidagriExcel(midagriStream);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error leyendo archivo MIDAGRI: {ex.Message}", ex);
        }

        List<Siniestro> siniestros;
        try
        {
            siniestros = _excelReader.ReadSiniestrosExcel(siniestrosStream);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error leyendo archivo Siniestros: {ex.Message}", ex);
        }

        // 2. Load static materia asegurada
        List<MateriaAsegurada> materia;
        try
        {
            materia = _staticData.LoadMateriaAsegurada();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error leyendo datos estáticos (Materia Asegurada): {ex.Message}", ex);
        }

        // 3. Normalize TIPO_SINIESTRO on both datasets
        foreach (var s in midagri)
            s.TipoSiniestro = NormalizeTipoSiniestro(s.TipoSiniestro);
        foreach (var s in siniestros)
            s.TipoSiniestro = NormalizeTipoSiniestro(s.TipoSiniestro);

        // 4. Combine datasets
        var combined = new List<Siniestro>(midagri.Count + siniestros.Count);
        combined.AddRange(midagri);
        combined.AddRange(siniestros);

        // 5. National metrics
        int totalAvisos = combined.Count;
        int totalAjustados = combined.Count(s =>
            string.Equals(s.EstadoInspeccion?.Trim(), EstadoInspeccionConstants.Cerrado, StringComparison.OrdinalIgnoreCase));
        double pctAjustados = totalAvisos > 0 ? (double)totalAjustados / totalAvisos * 100.0 : 0.0;

        double haIndemnizadas = combined.Sum(s => s.SupIndemnizada);
        double montoIndemnizado = combined.Sum(s => s.Indemnizacion);
        double montoDesembolsado = combined.Sum(s => s.MontoDesembolsado);
        int productoresDesembolso = (int)combined.Sum(s => s.NProductores);

        double primaTotal = materia.Sum(m => m.PrimaTotal);
        double primaNeta = materia.Sum(m => m.PrimaNeta);
        double supAsegurada = materia.Sum(m => m.SuperficieAsegurada);
        int prodAsegurados = (int)materia.Sum(m => m.ProductoresAsegurados);

        double indiceSiniestralidad = primaNeta > 0 ? montoIndemnizado / primaNeta * 100.0 : 0.0;
        double pctDesembolso = montoIndemnizado > 0 ? montoDesembolsado / montoIndemnizado * 100.0 : 0.0;

        int deptosConDesembolso = combined
            .Where(s => s.MontoDesembolsado > 0)
            .Select(s => (s.Departamento ?? string.Empty).Trim().ToUpper())
            .Distinct()
            .Count();

        // Empresas: group materia by empresa, count distinct departments
        var empresas = materia
            .GroupBy(m => (m.EmpresaAseguradora ?? string.Empty).Trim())
            .Where(g => !string.IsNullOrWhiteSpace(g.Key))
            .ToDictionary(
                g => g.Key,
                g => g.Select(m => (m.Departamento ?? string.Empty).Trim().ToUpper()).Distinct().Count()
            );

        // Empresas text
        var empresasTextParts = empresas
            .OrderBy(e => e.Key)
            .Select(e => $"{e.Key} ({e.Value} departamentos)");
        string empresasText = string.Join("; ", empresasTextParts);

        // 6. Cuadro 1: from materia, group by Departamento
        var cuadro1 = materia
            .GroupBy(m => (m.Departamento ?? string.Empty).Trim().ToUpper())
            .Where(g => !string.IsNullOrWhiteSpace(g.Key))
            .Select(g => new Cuadro1Row
            {
                Departamento = g.Key,
                PrimaTotal = g.Sum(m => m.PrimaTotal),
                Hectareas = g.Sum(m => m.SuperficieAsegurada),
                SumaAsegurada = g.Sum(m => m.ValoresAsegurados)
            })
            .OrderBy(r => r.Departamento)
            .ToList();

        // 7. Cuadro 2: from combined, group by Departamento
        var cuadro2 = combined
            .GroupBy(s => (s.Departamento ?? string.Empty).Trim().ToUpper())
            .Where(g => !string.IsNullOrWhiteSpace(g.Key))
            .Select(g => new Cuadro2Row
            {
                Departamento = g.Key,
                HaIndemnizadas = g.Sum(s => s.SupIndemnizada),
                MontoIndemnizado = g.Sum(s => s.Indemnizacion),
                MontoDesembolsado = g.Sum(s => s.MontoDesembolsado),
                Productores = g.Sum(s => s.NProductores)
            })
            .OrderBy(r => r.Departamento)
            .ToList();

        // 8. Cuadro 3: lluvia events only
        var lluviaRecords = combined
            .Where(s => TipoSiniestroConstants.LluviaTypes.Contains(
                NormalizeTipoSiniestro(s.TipoSiniestro)))
            .ToList();

        var cuadro3 = lluviaRecords
            .GroupBy(s => (s.Departamento ?? string.Empty).Trim().ToUpper())
            .Where(g => !string.IsNullOrWhiteSpace(g.Key))
            .Select(g => new Cuadro3Row
            {
                Departamento = g.Key,
                Avisos = g.Count(),
                HaIndemn = g.Sum(s => s.SupIndemnizada),
                MontoIndemnizado = g.Sum(s => s.Indemnizacion),
                MontoDesembolsado = g.Sum(s => s.MontoDesembolsado),
                Productores = g.Sum(s => s.NProductores)
            })
            .OrderBy(r => r.Departamento)
            .ToList();

        int totalLluvia = lluviaRecords.Count;
        double pctLluvia = totalAvisos > 0 ? (double)totalLluvia / totalAvisos * 100.0 : 0.0;

        // Lluvia por tipo
        var lluviaPorTipo = lluviaRecords
            .GroupBy(s => s.TipoSiniestro ?? string.Empty)
            .ToDictionary(g => g.Key, g => g.Count());

        // 9. Siniestros por tipo (value_counts equivalent)
        var siniestrosPorTipo = combined
            .GroupBy(s => s.TipoSiniestro ?? string.Empty)
            .ToDictionary(g => g.Key, g => g.Count())
            .OrderByDescending(kv => kv.Value)
            .ToDictionary(kv => kv.Key, kv => kv.Value);

        var top3 = siniestrosPorTipo
            .Take(3)
            .ToDictionary(kv => kv.Key, kv => kv.Value);

        // 10. Sorted unique departments
        var departamentosList = combined
            .Select(s => (s.Departamento ?? string.Empty).Trim().ToUpper())
            .Where(d => !string.IsNullOrWhiteSpace(d))
            .Distinct()
            .OrderBy(d => d)
            .ToList();

        // Build result
        var resultado = new DatosNacionalesDto
        {
            FechaCorte = DateTime.Now.ToString("dd/MM/yyyy"),
            Metricas = new MetricasDto
            {
                TotalAvisos = totalAvisos,
                TotalAjustados = totalAjustados,
                PctAjustados = Math.Round(pctAjustados, 2),
                HaIndemnizadas = Math.Round(haIndemnizadas, 2),
                MontoIndemnizado = Math.Round(montoIndemnizado, 2),
                MontoDesembolsado = Math.Round(montoDesembolsado, 2),
                ProductoresDesembolso = productoresDesembolso,
                PrimaTotal = Math.Round(primaTotal, 2),
                PrimaNeta = Math.Round(primaNeta, 2),
                SupAsegurada = Math.Round(supAsegurada, 2),
                ProdAsegurados = prodAsegurados,
                IndiceSiniestralidad = Math.Round(indiceSiniestralidad, 2),
                PctDesembolso = Math.Round(pctDesembolso, 2),
                DeptosConDesembolso = deptosConDesembolso
            },
            EmpresasText = empresasText,
            Empresas = empresas,
            Cuadro1 = cuadro1,
            Cuadro2 = cuadro2,
            Cuadro3 = cuadro3,
            TotalLluvia = totalLluvia,
            PctLluvia = Math.Round(pctLluvia, 2),
            LluviaPorTipo = lluviaPorTipo,
            SiniestrosPorTipo = siniestrosPorTipo,
            Top3Siniestros = top3,
            DepartamentosList = departamentosList,
            Midagri = midagri,
            SiniestrosOriginal = siniestros,
            Materia = materia
        };

        return resultado;
    }

    // ────────────────────────────────────────────
    //  PUBLIC: GetDepartamentoData
    // ────────────────────────────────────────────
    public DatosDepartamentalDto GetDepartamentoData(DatosNacionalesDto datos, string departamento)
    {
        string deptoUpper = (departamento ?? string.Empty).Trim().ToUpper();

        // 1. Filter combined data by department
        var combined = new List<Siniestro>(datos.Midagri.Count + datos.SiniestrosOriginal.Count);
        combined.AddRange(datos.Midagri);
        combined.AddRange(datos.SiniestrosOriginal);

        var deptoData = combined
            .Where(s => string.Equals(
                (s.Departamento ?? string.Empty).Trim(),
                deptoUpper,
                StringComparison.OrdinalIgnoreCase))
            .ToList();

        // 2. Get materia data for department
        var materiaDepto = datos.Materia
            .Where(m => string.Equals(
                (m.Departamento ?? string.Empty).Trim(),
                deptoUpper,
                StringComparison.OrdinalIgnoreCase))
            .ToList();

        string empresa = materiaDepto.FirstOrDefault()?.EmpresaAseguradora ?? "N/D";
        double primaNeta = materiaDepto.Sum(m => m.PrimaNeta);
        double supAsegurada = materiaDepto.Sum(m => m.SuperficieAsegurada);

        // 3. Department-specific metrics
        int totalAvisos = deptoData.Count;
        double haIndemnizadas = deptoData.Sum(s => s.SupIndemnizada);
        double montoIndemnizado = deptoData.Sum(s => s.Indemnizacion);
        double montoDesembolsado = deptoData.Sum(s => s.MontoDesembolsado);
        int productoresDesembolso = (int)deptoData.Sum(s => s.NProductores);

        // 4. Count indemnizables / no_indemnizables using Dictamen
        int indemnizables = deptoData.Count(s =>
            string.Equals((s.Dictamen ?? string.Empty).Trim(),
                DictamenConstants.Indemnizable,
                StringComparison.OrdinalIgnoreCase));
        int noIndemnizables = deptoData.Count(s =>
            string.Equals((s.Dictamen ?? string.Empty).Trim(),
                DictamenConstants.NoIndemnizable,
                StringComparison.OrdinalIgnoreCase));

        // Estados (EstadoInspeccion value counts)
        var estados = deptoData
            .GroupBy(s => (s.EstadoInspeccion ?? string.Empty).Trim().ToUpper())
            .Where(g => !string.IsNullOrWhiteSpace(g.Key))
            .ToDictionary(g => g.Key, g => g.Count());

        // 5. Avisos por tipo: TipoSiniestro value counts with percentages
        var avisosTipoDict = deptoData
            .GroupBy(s => s.TipoSiniestro ?? string.Empty)
            .Where(g => !string.IsNullOrWhiteSpace(g.Key))
            .OrderByDescending(g => g.Count())
            .ToDictionary(g => g.Key, g => g.Count());

        var avisosTipo = avisosTipoDict
            .Select(kv =>
            {
                double pct = totalAvisos > 0 ? (double)kv.Value / totalAvisos * 100.0 : 0.0;
                return new string[] { kv.Key, kv.Value.ToString(), $"{pct:F1}%" };
            })
            .ToList();

        // 6. Distribucion por provincia
        var distProvincia = deptoData
            .GroupBy(s => (s.Provincia ?? string.Empty).Trim().ToUpper())
            .Where(g => !string.IsNullOrWhiteSpace(g.Key))
            .OrderBy(g => g.Key)
            .Select(g =>
            {
                int avisos = g.Count();
                double supIndemn = g.Sum(s => s.SupIndemnizada);
                double productores = g.Sum(s => s.NProductores);
                double indemniz = g.Sum(s => s.Indemnizacion);
                double desembolso = g.Sum(s => s.MontoDesembolsado);
                double pctAvance = indemniz > 0 ? desembolso / indemniz * 100.0 : 0.0;

                return new string[]
                {
                    g.Key,
                    avisos.ToString(),
                    FmtNum(supIndemn),
                    FmtNum(productores, 0),
                    FmtNum(indemniz),
                    FmtNum(desembolso),
                    $"{pctAvance:F1}%"
                };
            })
            .ToList();

        // 7. Eventos recientes: last 30 days by FechaAviso/FechaSiniestro, top 20
        DateTime cutoff = DateTime.Now.AddDays(-30);

        var eventosRecientes = deptoData
            .Where(s =>
            {
                DateTime? refDate = s.FechaAviso ?? s.FechaSiniestro;
                return refDate.HasValue && refDate.Value >= cutoff;
            })
            .OrderByDescending(s => s.FechaAviso ?? s.FechaSiniestro)
            .Take(20)
            .Select(s =>
            {
                string fecha = (s.FechaAviso ?? s.FechaSiniestro)?.ToString("dd/MM/yyyy") ?? "";
                string provincia = (s.Provincia ?? string.Empty).Trim();
                string distritoSector = string.IsNullOrWhiteSpace(s.SectorEstadistico)
                    ? (s.Distrito ?? string.Empty).Trim()
                    : $"{(s.Distrito ?? string.Empty).Trim()} / {s.SectorEstadistico.Trim()}";
                string cultivo = (s.TipoCultivo ?? string.Empty).Trim();
                string estado = (s.EstadoInspeccion ?? string.Empty).Trim();

                return new string[] { fecha, provincia, distritoSector, cultivo, estado };
            })
            .ToList();

        // 8. Build resumen_operativo text
        double indiceSiniestralidad = primaNeta > 0 ? montoIndemnizado / primaNeta * 100.0 : 0.0;
        double pctDesembolso = montoIndemnizado > 0 ? montoDesembolsado / montoIndemnizado * 100.0 : 0.0;

        string resumenOperativo = BuildResumenOperativo(
            deptoUpper, empresa, totalAvisos, haIndemnizadas, montoIndemnizado,
            primaNeta, indiceSiniestralidad, indemnizables, noIndemnizables);

        string resumenDesembolso = BuildResumenDesembolso(
            deptoUpper, montoDesembolsado, montoIndemnizado, pctDesembolso, productoresDesembolso);

        return new DatosDepartamentalDto
        {
            Departamento = deptoUpper,
            Empresa = empresa,
            PrimaNeta = Math.Round(primaNeta, 2),
            SupAsegurada = Math.Round(supAsegurada, 2),
            TotalAvisos = totalAvisos,
            HaIndemnizadas = Math.Round(haIndemnizadas, 2),
            MontoIndemnizado = Math.Round(montoIndemnizado, 2),
            MontoDesembolsado = Math.Round(montoDesembolsado, 2),
            ProductoresDesembolso = productoresDesembolso,
            Indemnizables = indemnizables,
            NoIndemnizables = noIndemnizables,
            FechaCorte = DateTime.Now.ToString("dd/MM/yyyy"),
            Estados = estados,
            AvisosTipo = avisosTipo,
            DistProvincia = distProvincia,
            EventosRecientes = eventosRecientes,
            ResumenOperativo = resumenOperativo,
            ResumenDesembolso = resumenDesembolso
        };
    }

    // ────────────────────────────────────────────
    //  PRIVATE: NormalizeTipoSiniestro
    // ────────────────────────────────────────────
    private static string NormalizeTipoSiniestro(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return string.Empty;
        var normalized = s.Normalize(NormalizationForm.FormKD);
        var sb = new StringBuilder();
        foreach (var c in normalized)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                sb.Append(c);
        }
        return sb.ToString().Trim().ToUpper();
    }

    // ────────────────────────────────────────────
    //  PUBLIC STATIC: FmtNum  (Spanish locale)
    // ────────────────────────────────────────────
    public static string FmtNum(double val, int dec = 2)
    {
        if (dec == 0)
            return val.ToString("N0", new CultureInfo("es-PE"));
        return val.ToString($"N{dec}", new CultureInfo("es-PE"));
    }

    // ────────────────────────────────────────────
    //  PRIVATE: BuildResumenOperativo
    // ────────────────────────────────────────────
    private static string BuildResumenOperativo(
        string depto, string empresa, int totalAvisos,
        double haIndemnizadas, double montoIndemnizado,
        double primaNeta, double indiceSiniestralidad,
        int indemnizables, int noIndemnizables)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"En el departamento de {depto}, asegurado por {empresa}, " +
            $"se han registrado un total de {FmtNum(totalAvisos, 0)} avisos de siniestro.");
        sb.AppendLine($"Se han indemnizado {FmtNum(haIndemnizadas)} hectáreas por un monto de " +
            $"S/ {FmtNum(montoIndemnizado)}.");
        sb.AppendLine($"La prima neta asignada al departamento es de S/ {FmtNum(primaNeta)}, " +
            $"lo que arroja un índice de siniestralidad de {FmtNum(indiceSiniestralidad)}%.");
        sb.AppendLine($"Del total de avisos evaluados, {FmtNum(indemnizables, 0)} fueron dictaminados " +
            $"como INDEMNIZABLES y {FmtNum(noIndemnizables, 0)} como NO INDEMNIZABLES.");
        return sb.ToString().TrimEnd();
    }

    // ────────────────────────────────────────────
    //  PRIVATE: BuildResumenDesembolso
    // ────────────────────────────────────────────
    private static string BuildResumenDesembolso(
        string depto, double montoDesembolsado,
        double montoIndemnizado, double pctDesembolso,
        int productoresDesembolso)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"En {depto}, se ha desembolsado S/ {FmtNum(montoDesembolsado)} de un total " +
            $"indemnizado de S/ {FmtNum(montoIndemnizado)}, lo que representa un avance del " +
            $"{FmtNum(pctDesembolso)}%.");
        sb.AppendLine($"Este desembolso ha beneficiado a {FmtNum(productoresDesembolso, 0)} productores.");
        return sb.ToString().TrimEnd();
    }
}
