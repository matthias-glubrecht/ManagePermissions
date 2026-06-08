namespace ManagePermissions.Options;

/// <summary>
/// Konfiguration zur Validierung des Aufrufer-Tokens (Entra ID).
/// App-Settings-Präfix: <c>AzureAd__</c> (z. B. <c>AzureAd__TenantId</c>).
/// </summary>
public sealed class EntraAuthOptions
{
    public const string SectionName = "AzureAd";

    /// <summary>Verzeichnis-(Mandanten-)ID, gegen die das Token validiert wird.</summary>
    public string TenantId { get; set; } = string.Empty;

    /// <summary>Client-ID der API-App-Registrierung dieser Function (= erwartete Audience).</summary>
    public string ClientId { get; set; } = string.Empty;

    /// <summary>Delegierter Scope, den das Token enthalten muss.</summary>
    public string RequiredScope { get; set; } = "access_as_user";

    /// <summary>Objekt-ID der Sicherheitsgruppe, deren Mitglieder aufrufen dürfen.</summary>
    public string AllowedGroupId { get; set; } = string.Empty;
}
