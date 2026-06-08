namespace ManagePermissions.Models;

/// <summary>
/// Eingabevertrag des <c>ManagePermissions</c>-Endpunkts.
/// </summary>
public sealed class ManagePermissionsRequest
{
    /// <summary>Auszuführende Aktion: <c>grant</c> oder <c>reset</c>.</summary>
    public string? Action { get; set; }

    /// <summary>Absolute URL des SharePoint-Webs, z. B. <c>https://contoso.sharepoint.com/sites/team</c>.</summary>
    public string? WebUrl { get; set; }

    /// <summary>GUID der Liste innerhalb des Webs.</summary>
    public string? ListId { get; set; }

    /// <summary>Numerische ID des Listenelements.</summary>
    public int? ItemId { get; set; }

    /// <summary>UPN des Benutzers (z. B. <c>user@domain.com</c>); nur für <c>grant</c>.</summary>
    public string? UserPrincipalName { get; set; }

    /// <summary>
    /// Zu setzende Berechtigungsstufe; nur für <c>grant</c>.
    /// Erlaubt: <c>Read</c>, <c>Contribute</c>, <c>Edit</c>, <c>Design</c>, <c>FullControl</c>.
    /// </summary>
    public string? PermissionLevel { get; set; }
}
