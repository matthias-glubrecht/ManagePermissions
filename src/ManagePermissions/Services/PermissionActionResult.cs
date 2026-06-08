namespace ManagePermissions.Services;

/// <summary>
/// Ergebnis einer Berechtigungsaktion (grant/reset) auf einem Listenelement.
/// </summary>
public sealed record PermissionActionResult(bool Succeeded, int StatusCode, string Message)
{
    public static PermissionActionResult Ok(string message) => new(true, 200, message);

    public static PermissionActionResult Fail(int statusCode, string message) => new(false, statusCode, message);
}
