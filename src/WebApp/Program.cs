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

app.Run();
