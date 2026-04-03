namespace Infrastructure.AutoDownload;

using Application.Contracts.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Playwright;

/// <summary>
/// Downloads SAC data from SISGAQSAC (Rimac) and Agroevaluaciones (La Positiva) portals
/// using headless Chromium via Playwright.
/// Credentials from appsettings.json or environment variables.
/// </summary>
public class AutoDownloadService : IAutoDownloadService
{
    private readonly IConfiguration _config;

    public AutoDownloadService(IConfiguration config) => _config = config;

    public bool IsConfigured()
    {
        var (rimacEmail, rimacPass) = GetRimacCredentials();
        var (lpUser, lpPass) = GetLaPositivaCredentials();
        return !string.IsNullOrEmpty(rimacEmail) && !string.IsNullOrEmpty(rimacPass)
            && !string.IsNullOrEmpty(lpUser) && !string.IsNullOrEmpty(lpPass);
    }

    private (string? email, string? password) GetRimacCredentials()
    {
        var email = _config["AutoDownload:Rimac:Email"]
            ?? Environment.GetEnvironmentVariable("RIMAC_EMAIL");
        var pass = _config["AutoDownload:Rimac:Password"]
            ?? Environment.GetEnvironmentVariable("RIMAC_PASSWORD");
        return (email, pass);
    }

    private (string? user, string? password) GetLaPositivaCredentials()
    {
        var user = _config["AutoDownload:LaPositiva:Usuario"]
            ?? Environment.GetEnvironmentVariable("LP_USUARIO");
        var pass = _config["AutoDownload:LaPositiva:Password"]
            ?? Environment.GetEnvironmentVariable("LP_PASSWORD");
        return (user, pass);
    }

    public async Task<AutoDownloadResult> DescargarRimacAsync(Action<string>? onProgress = null)
    {
        var (email, password) = GetRimacCredentials();
        if (string.IsNullOrEmpty(email) || string.IsNullOrEmpty(password))
            return new AutoDownloadResult { Error = "Credenciales de Rimac no configuradas." };

        try
        {
            onProgress?.Invoke("Iniciando navegador...");
            using var playwright = await Playwright.CreateAsync();
            await using var browser = await playwright.Chromium.LaunchAsync(new()
            {
                Headless = true,
                ExecutablePath = Environment.GetEnvironmentVariable("PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH")
            });
            var page = await browser.NewPageAsync();

            // 1. Navigate to login
            onProgress?.Invoke("Navegando a SISGAQSAC...");
            await page.GotoAsync("https://www.sisgaqsac.pe", new() { Timeout = 30000 });
            await page.WaitForLoadStateAsync(LoadState.NetworkIdle);

            // 2. Fill credentials
            onProgress?.Invoke("Ingresando credenciales...");
            await page.FillAsync("input[type='email']", email);
            await page.FillAsync("input[type='password']", password);

            // 3. Click login
            await page.ClickAsync("button:has-text('Acceder')");
            await page.WaitForLoadStateAsync(LoadState.NetworkIdle);

            // Verify login
            var salir = await page.QuerySelectorAsync("text=Salir");
            if (salir == null)
                return new AutoDownloadResult { Error = "Login fallido en SISGAQSAC." };

            // 4. Click Excel download
            onProgress?.Invoke("Descargando Excel...");
            var download = await page.RunAndWaitForDownloadAsync(async () =>
            {
                await page.ClickAsync("button:has-text('Excel')");
            }, new() { Timeout = 60000 });

            var tempPath = Path.GetTempFileName() + ".xlsx";
            await download.SaveAsAsync(tempPath);
            var bytes = await File.ReadAllBytesAsync(tempPath);
            File.Delete(tempPath);

            onProgress?.Invoke("Descarga Rimac completada.");
            return new AutoDownloadResult
            {
                Success = true, FileBytes = bytes,
                FileName = $"rimac_siniestros_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
            };
        }
        catch (Exception ex)
        {
            return new AutoDownloadResult { Error = $"Error Rimac: {ex.Message}" };
        }
    }

    public async Task<AutoDownloadResult> DescargarLaPositivaAsync(Action<string>? onProgress = null)
    {
        var (usuario, password) = GetLaPositivaCredentials();
        if (string.IsNullOrEmpty(usuario) || string.IsNullOrEmpty(password))
            return new AutoDownloadResult { Error = "Credenciales de La Positiva no configuradas." };

        try
        {
            onProgress?.Invoke("Iniciando navegador...");
            using var playwright = await Playwright.CreateAsync();
            await using var browser = await playwright.Chromium.LaunchAsync(new()
            {
                Headless = true,
                ExecutablePath = Environment.GetEnvironmentVariable("PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH")
            });
            var page = await browser.NewPageAsync();

            // 1. Navigate to login
            onProgress?.Invoke("Navegando a Agroevaluaciones...");
            await page.GotoAsync("https://catastrofico.agroevaluaciones.com/login", new() { Timeout = 30000 });
            await page.WaitForLoadStateAsync(LoadState.NetworkIdle);

            // 2. Fill credentials
            onProgress?.Invoke("Ingresando credenciales...");
            await page.FillAsync("input[type='text']", usuario);
            await page.FillAsync("input[type='password']", password);

            // 3. Click login
            await page.ClickAsync("button:has-text('Iniciar')");
            await page.WaitForLoadStateAsync(LoadState.NetworkIdle);
            await page.WaitForTimeoutAsync(2000);

            // 4. Navigate to Avisos > Todos
            onProgress?.Invoke("Navegando a Avisos > Todos...");
            await page.ClickAsync("text=Avisos");
            await page.WaitForTimeoutAsync(1000);
            await page.ClickAsync("text=Todos");
            await page.WaitForLoadStateAsync(LoadState.NetworkIdle);
            await page.WaitForTimeoutAsync(2000);

            // 5. Find Midagri button (exclude FOGASA)
            onProgress?.Invoke("Buscando boton Midagri...");
            var buttons = await page.QuerySelectorAllAsync("button:has-text('Midagri')");
            IElementHandle? midagriBtn = null;
            foreach (var btn in buttons)
            {
                var text = await btn.InnerTextAsync();
                if (!text.Contains("FOGASA", StringComparison.OrdinalIgnoreCase))
                {
                    midagriBtn = btn;
                    break;
                }
            }

            if (midagriBtn == null)
                return new AutoDownloadResult { Error = "No se encontro el boton Midagri." };

            // 6. Download with long timeout (~70s processing)
            onProgress?.Invoke("Descargando MIDAGRI (puede tomar ~70s)...");
            var download = await page.RunAndWaitForDownloadAsync(async () =>
            {
                await midagriBtn.ClickAsync();
            }, new() { Timeout = 300000 }); // 5 minutes

            var tempPath = Path.GetTempFileName() + ".xlsx";
            await download.SaveAsAsync(tempPath);
            var bytes = await File.ReadAllBytesAsync(tempPath);
            File.Delete(tempPath);

            onProgress?.Invoke("Descarga La Positiva completada.");
            return new AutoDownloadResult
            {
                Success = true, FileBytes = bytes,
                FileName = $"lp_midagri_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
            };
        }
        catch (Exception ex)
        {
            return new AutoDownloadResult { Error = $"Error La Positiva: {ex.Message}" };
        }
    }
}
