namespace ManagePermissions.Options;

/// <summary>
/// Konfiguration für den SharePoint-Zugriff.
/// App-Settings-Präfix: <c>SharePoint__</c> (z. B. <c>SharePoint__AllowedHosts</c>).
/// </summary>
public sealed class SharePointOptions
{
    public const string SectionName = "SharePoint";

    /// <summary>
    /// Komma-getrennte Allowlist erlaubter SharePoint-Hostnamen (z. B.
    /// <c>contoso.sharepoint.com</c>). Schützt vor missbräuchlichen <c>webUrl</c>-Zielen (SSRF).
    /// Bleibt der Wert leer, werden alle <c>*.sharepoint.com</c>-Hosts akzeptiert.
    /// </summary>
    public string AllowedHosts { get; set; } = string.Empty;

    /// <summary>Geparste Allowlist (leere Einträge entfernt, getrimmt).</summary>
    public IReadOnlyList<string> AllowedHostsList =>
        AllowedHosts.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
}
