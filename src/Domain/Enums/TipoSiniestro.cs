namespace Domain.Enums;

public static class TipoSiniestroConstants
{
    public static readonly HashSet<string> LluviaTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "INUNDACION", "INUNDACIÓN", "HUAYCO", "LLUVIAS EXCESIVAS", "DESLIZAMIENTO", "DESLIZAMIENTOS"
    };
}
