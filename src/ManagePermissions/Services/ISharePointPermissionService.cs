using ManagePermissions.Models;

namespace ManagePermissions.Services;

/// <summary>
/// Setzt bzw. entfernt Berechtigungen auf SharePoint-Listenelementen.
/// </summary>
public interface ISharePointPermissionService
{
    /// <summary>Vergibt einem Benutzer eine Berechtigungsstufe auf einem Listenelement.</summary>
    Task<PermissionActionResult> GrantAsync(ManagePermissionsRequest request, string callerUpn, CancellationToken ct);

    /// <summary>Stellt die Vererbung eines Listenelements wieder her (entfernt eindeutige Berechtigungen).</summary>
    Task<PermissionActionResult> ResetAsync(ManagePermissionsRequest request, string callerUpn, CancellationToken ct);
}
