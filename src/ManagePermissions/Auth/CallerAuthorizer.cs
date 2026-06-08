using System.Security.Claims;
using ManagePermissions.Options;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.IdentityModel.JsonWebTokens;
using Microsoft.IdentityModel.Protocols;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using Microsoft.IdentityModel.Tokens;

namespace ManagePermissions.Auth;

/// <summary>
/// Validiert das vom SPFx-Web-Part mitgesendete Entra-ID-Bearer-Token im Code:
/// Signatur/Issuer/Audience/Lifetime, den erforderlichen Scope und die Gruppenmitgliedschaft.
/// </summary>
public sealed class CallerAuthorizer
{
    private const string BearerPrefix = "Bearer ";

    private readonly EntraAuthOptions _options;
    private readonly ILogger<CallerAuthorizer> _logger;
    private readonly ConfigurationManager<OpenIdConnectConfiguration> _configManager;

    public CallerAuthorizer(IOptions<EntraAuthOptions> options, ILogger<CallerAuthorizer> logger)
    {
        _options = options.Value;
        _logger = logger;

        var metadataAddress =
            $"https://login.microsoftonline.com/{_options.TenantId}/v2.0/.well-known/openid-configuration";
        _configManager = new ConfigurationManager<OpenIdConnectConfiguration>(
            metadataAddress,
            new OpenIdConnectConfigurationRetriever(),
            new HttpDocumentRetriever());
    }

    public async Task<CallerResult> AuthorizeAsync(string? authorizationHeader, CancellationToken ct)
    {
        if (string.IsNullOrEmpty(_options.TenantId) || string.IsNullOrEmpty(_options.ClientId))
        {
            _logger.LogError("AzureAd:TenantId/ClientId ist nicht konfiguriert.");
            return CallerResult.Fail(500, "Server-Authentifizierung ist nicht konfiguriert.");
        }

        if (string.IsNullOrWhiteSpace(authorizationHeader) ||
            !authorizationHeader.StartsWith(BearerPrefix, StringComparison.OrdinalIgnoreCase))
        {
            return CallerResult.Fail(401, "Kein Bearer-Token im Authorization-Header.");
        }

        var token = authorizationHeader[BearerPrefix.Length..].Trim();

        OpenIdConnectConfiguration config;
        try
        {
            config = await _configManager.GetConfigurationAsync(ct);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "OpenID-Connect-Metadaten konnten nicht geladen werden.");
            return CallerResult.Fail(500, "Token-Konfiguration ist nicht verfügbar.");
        }

        var parameters = new TokenValidationParameters
        {
            ValidateIssuer = true,
            ValidIssuer = config.Issuer,
            ValidateAudience = true,
            ValidAudiences = new[] { _options.ClientId, $"api://{_options.ClientId}" },
            ValidateLifetime = true,
            ValidateIssuerSigningKey = true,
            IssuerSigningKeys = config.SigningKeys,
            ClockSkew = TimeSpan.FromMinutes(2),
        };

        var handler = new JsonWebTokenHandler();
        var validation = await handler.ValidateTokenAsync(token, parameters);
        if (!validation.IsValid)
        {
            _logger.LogWarning(validation.Exception, "Token-Validierung fehlgeschlagen.");
            return CallerResult.Fail(401, "Token ist ungültig oder abgelaufen.");
        }

        var identity = validation.ClaimsIdentity;

        // Erforderlichen delegierten Scope (scp) prüfen.
        var scopeClaim = identity.FindFirst("scp")?.Value
                         ?? identity.FindFirst("http://schemas.microsoft.com/identity/claims/scope")?.Value;
        var scopes = (scopeClaim ?? string.Empty).Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (!scopes.Contains(_options.RequiredScope, StringComparer.OrdinalIgnoreCase))
        {
            return CallerResult.Fail(403, $"Erforderlicher Scope '{_options.RequiredScope}' fehlt im Token.");
        }

        // Gruppenmitgliedschaft prüfen (sofern konfiguriert).
        if (!string.IsNullOrEmpty(_options.AllowedGroupId))
        {
            var groups = identity.FindAll("groups").Select(c => c.Value);
            if (!groups.Contains(_options.AllowedGroupId, StringComparer.OrdinalIgnoreCase))
            {
                var hasGroupOverage = identity.FindFirst("_claim_names")?.Value?
                    .Contains("groups", StringComparison.OrdinalIgnoreCase) == true;
                if (hasGroupOverage)
                {
                    _logger.LogWarning(
                        "Group-Overage im Token – Gruppen nicht enthalten. groupMembershipClaims=ApplicationGroup setzen.");
                    return CallerResult.Fail(403,
                        "Gruppenmitgliedschaft konnte nicht aus dem Token ermittelt werden (Group-Overage).");
                }

                return CallerResult.Fail(403, "Aufrufer ist nicht Mitglied der berechtigten Gruppe.");
            }
        }

        var callerUpn = identity.FindFirst("preferred_username")?.Value
                        ?? identity.FindFirst(ClaimTypes.Upn)?.Value
                        ?? identity.FindFirst("upn")?.Value
                        ?? "(unbekannt)";
        return CallerResult.Success(callerUpn);
    }
}
