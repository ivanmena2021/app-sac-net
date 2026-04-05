using WebApp.Components;
using Application.Contracts.Services;
using Application.Services;
using Infrastructure;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

// Register Clean Architecture services
builder.Services.AddScoped<IDataProcessorService, DataProcessorService>();
builder.Services.AddInfrastructure();
builder.Services.AddHttpClient();

// Increase SignalR message size for large file uploads
builder.Services.AddSignalR(o =>
{
    o.MaximumReceiveMessageSize = 50 * 1024 * 1024; // 50 MB
});

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    app.UseHsts();
}

app.UseStaticFiles();
app.UseAntiforgery();

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

// Health check: verify Python API is reachable at startup
_ = Task.Run(async () =>
{
    var pythonUrl = builder.Configuration["PythonApi:BaseUrl"] ?? "http://localhost:8000";
    using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
    for (int attempt = 1; attempt <= 5; attempt++)
    {
        try
        {
            var resp = await client.GetAsync($"{pythonUrl}/health");
            if (resp.IsSuccessStatusCode)
            {
                Console.WriteLine($"[SAC] Python API OK at {pythonUrl}");
                return;
            }
        }
        catch { /* retry */ }
        Console.WriteLine($"[SAC] Python API not reachable at {pythonUrl} (attempt {attempt}/5)");
        await Task.Delay(5000);
    }
    Console.WriteLine($"[SAC] WARNING: Python API at {pythonUrl} not reachable after 5 attempts. Document generation will fail.");
});

app.Run();
