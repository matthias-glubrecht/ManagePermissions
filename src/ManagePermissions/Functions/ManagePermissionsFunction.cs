using ManagePermissions.Auth;
using ManagePermissions.Models;
using ManagePermissions.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace ManagePermissions.Functions;

/// <summary>
/// HTTP-Endpunkt zum Setzen/Entfernen von Berechtigungen auf SharePoint-Listenelementen.
/// Aufruf aus einem SPFx-Web-Part mit einem Entra-ID-Bearer-Token (Scope <c>access_as_user</c>).
/// </summary>
public sealed class ManagePermissionsFunction
{
    private readonly CallerAuthorizer _authorizer;
    private readonly ISharePointPermissionService _permissionService;
    private readonly ILogger<ManagePermissionsFunction> _logger;

    public ManagePermissionsFunction(
        CallerAuthorizer authorizer,
        ISharePointPermissionService permissionService,
        ILogger<ManagePermissionsFunction> logger)
    {
        _authorizer = authorizer;
        _permissionService = permissionService;
        _logger = logger;
    }

    [Function("ManagePermissions")]
    public async Task<IActionResult> Run(
        [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "ManagePermissions")] HttpRequest req,
        CancellationToken ct)
    {
        // 1) Aufrufer authentifizieren + autorisieren (JWT-Validierung im Code).
        var auth = await _authorizer.AuthorizeAsync(req.Headers.Authorization.ToString(), ct);
        if (!auth.Succeeded)
        {
            return Json(auth.StatusCode, new { error = auth.Message });
        }

        // 2) Anfrage-Body als JSON lesen.
        ManagePermissionsRequest? body;
        try
        {
            body = await req.ReadFromJsonAsync<ManagePermissionsRequest>(ct);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Body konnte nicht als JSON gelesen werden.");
            return Json(400, new { error = "Ungültiger JSON-Body." });
        }

        if (body is null || string.IsNullOrWhiteSpace(body.Action))
        {
            return Json(400, new { error = "Feld 'action' ist erforderlich ('grant' oder 'reset')." });
        }

        // 3) Aktion ausführen.
        var result = body.Action.Trim().ToLowerInvariant() switch
        {
            "grant" => await _permissionService.GrantAsync(body, auth.CallerUpn!, ct),
            "reset" => await _permissionService.ResetAsync(body, auth.CallerUpn!, ct),
            _ => PermissionActionResult.Fail(400, "Unbekannte 'action'. Erlaubt: 'grant' oder 'reset'."),
        };

        if (result.Succeeded)
        {
            return Json(result.StatusCode, new { ok = true, message = result.Message });
        }

        return Json(result.StatusCode, new { error = result.Message });
    }

    private static IActionResult Json(int statusCode, object body) =>
        new ObjectResult(body) { StatusCode = statusCode };
}
