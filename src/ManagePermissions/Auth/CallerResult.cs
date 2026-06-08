namespace ManagePermissions.Auth;

/// <summary>
/// Ergebnis der Aufrufer-Validierung (Authentifizierung + Autorisierung).
/// </summary>
public sealed record CallerResult(bool Succeeded, int StatusCode, string? Message, string? CallerUpn)
{
    public static CallerResult Success(string callerUpn) => new(true, 200, null, callerUpn);

    public static CallerResult Fail(int statusCode, string message) => new(false, statusCode, message, null);
}
