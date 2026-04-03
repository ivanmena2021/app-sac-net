namespace Application.Contracts.Services;

using Application.DTOs.Responses;

public interface IQueryLlmService
{
    bool IsAvailable();
    Task<QueryLlmResult> ProcessQueryAsync(string question, DatosNacionalesDto datos);
}

public class QueryLlmResult
{
    public string? Prose { get; set; }
    public string? Sql { get; set; }
    public List<Dictionary<string, object>>? Data { get; set; }
    public Dictionary<string, object>? Summary { get; set; }
    public string? Error { get; set; }
}
