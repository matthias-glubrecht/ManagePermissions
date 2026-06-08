using Azure.Core;
using Azure.Identity;
using ManagePermissions.Auth;
using ManagePermissions.Options;
using ManagePermissions.Services;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Builder;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using PnP.Core.Auth;

var builder = FunctionsApplication.CreateBuilder(args);

builder.ConfigureFunctionsWebApplication();

builder.Services
    .AddApplicationInsightsTelemetryWorkerService()
    .ConfigureFunctionsApplicationInsights();

// --- Konfiguration ---
builder.Services.Configure<EntraAuthOptions>(builder.Configuration.GetSection(EntraAuthOptions.SectionName));
builder.Services.Configure<SharePointOptions>(builder.Configuration.GetSection(SharePointOptions.SectionName));

// --- Aufrufer-Validierung (Token-Prüfung im Code) ---
builder.Services.AddSingleton<CallerAuthorizer>();

// --- SharePoint-Zugriff über das PnP Core SDK ---
builder.Services.AddPnPCore(options =>
{
    // SharePoint-REST statt Graph nutzen, damit die Managed Identity nur das
    // SharePoint-Recht (Sites.Selected + FullControl) benötigt und keine Graph-Rechte.
    options.PnPContext.GraphFirst = false;
});

// Token-Beschaffung für SharePoint über die Managed Identity der Function App.
// DefaultAzureCredential = Managed Identity in Azure, 'az login' lokal.
builder.Services.AddSingleton<ExternalAuthenticationProvider>(_ =>
{
    var credential = new DefaultAzureCredential();
    return new ExternalAuthenticationProvider((resource, _) =>
    {
        var scope = $"https://{resource.Authority}/.default";
        var token = credential.GetToken(new TokenRequestContext(new[] { scope }), CancellationToken.None);
        return token.Token;
    });
});

builder.Services.AddSingleton<ISharePointPermissionService, SharePointPermissionService>();

builder.Build().Run();
