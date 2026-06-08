using ManagePermissions.Models;
using ManagePermissions.Options;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using PnP.Core.Auth;
using PnP.Core.Model.Security;
using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;
using PnP.Core.Services;

namespace ManagePermissions.Services;

/// <summary>
/// Implementiert das Setzen/Entfernen von Berechtigungen auf Listenelementen über das
/// PnP Core SDK. Der Zugriff auf SharePoint erfolgt app-only über die Managed Identity
/// der Function App.
/// </summary>
public sealed class SharePointPermissionService : ISharePointPermissionService
{
    /// <summary>Allowlist erlaubter Berechtigungsstufen → locale-unabhängiger <see cref="RoleType"/>.</summary>
    private static readonly IReadOnlyDictionary<string, RoleType> PermissionLevels =
        new Dictionary<string, RoleType>(StringComparer.OrdinalIgnoreCase)
        {
            ["Read"] = RoleType.Reader,
            ["Contribute"] = RoleType.Contributor,
            ["Edit"] = RoleType.Editor,
            ["Design"] = RoleType.WebDesigner,
            ["FullControl"] = RoleType.Administrator,
        };

    private readonly IPnPContextFactory _contextFactory;
    private readonly IAuthenticationProvider _authProvider;
    private readonly SharePointOptions _options;
    private readonly ILogger<SharePointPermissionService> _logger;

    public SharePointPermissionService(
        IPnPContextFactory contextFactory,
        ExternalAuthenticationProvider authProvider,
        IOptions<SharePointOptions> options,
        ILogger<SharePointPermissionService> logger)
    {
        _contextFactory = contextFactory;
        _authProvider = authProvider;
        _options = options.Value;
        _logger = logger;
    }

    public async Task<PermissionActionResult> GrantAsync(ManagePermissionsRequest request, string callerUpn, CancellationToken ct)
    {
        if (!TryValidateCommon(request, out var webUri, out var listId, out var commonError))
        {
            return commonError!;
        }

        if (string.IsNullOrWhiteSpace(request.UserPrincipalName))
        {
            return PermissionActionResult.Fail(400, "Feld 'userPrincipalName' ist erforderlich.");
        }

        if (string.IsNullOrWhiteSpace(request.PermissionLevel) ||
            !PermissionLevels.TryGetValue(request.PermissionLevel, out var roleType))
        {
            return PermissionActionResult.Fail(400,
                $"Ungültige 'permissionLevel'. Erlaubt: {string.Join(", ", PermissionLevels.Keys)}.");
        }

        try
        {
            using var context = await _contextFactory.CreateAsync(webUri, _authProvider);

            // Rollendefinition locale-unabhängig über den RoleType auflösen.
            var web = await context.Web.GetAsync(p => p.RoleDefinitions);
            var roleDefinition = web.RoleDefinitions.AsRequested().FirstOrDefault(r => r.RoleTypeKind == roleType);
            if (roleDefinition is null)
            {
                return PermissionActionResult.Fail(400,
                    $"Berechtigungsstufe '{request.PermissionLevel}' ist auf dieser Website nicht verfügbar.");
            }

            // Benutzer im User-Information-List des Webs sicherstellen.
            ISharePointUser user;
            try
            {
                user = await context.Web.EnsureUserAsync(request.UserPrincipalName!);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Benutzer '{Upn}' konnte nicht aufgelöst werden.", request.UserPrincipalName);
                return PermissionActionResult.Fail(404, $"Benutzer '{request.UserPrincipalName}' wurde nicht gefunden.");
            }

            var (item, lookupError) = await GetItemAsync(context, listId, request.ItemId!.Value);
            if (item is null)
            {
                return lookupError!;
            }

            // Eindeutige Berechtigungen sicherstellen: Vererbung trennen und bestehende
            // Zuweisungen übernehmen, damit vorhandene Zugriffe erhalten bleiben.
            if (!item.HasUniqueRoleAssignments)
            {
                await item.BreakRoleInheritanceAsync(copyRoleAssignments: true, clearSubscopes: false);
            }

            await item.AddRoleDefinitionAsync(user.Id, roleDefinition);

            _logger.LogInformation(
                "GRANT durch {Caller}: {Upn} → {Level} auf {Web} Liste {List} Element {Item}.",
                callerUpn, request.UserPrincipalName, request.PermissionLevel, webUri, listId, request.ItemId);

            return PermissionActionResult.Ok(
                $"'{request.PermissionLevel}' für {request.UserPrincipalName} auf Element {request.ItemId} gesetzt.");
        }
        catch (PnP.Core.ServiceException ex)
        {
            return MapServiceException(ex);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unerwarteter Fehler beim Setzen der Berechtigung.");
            return PermissionActionResult.Fail(500, "Unerwarteter Fehler beim Setzen der Berechtigung.");
        }
    }

    public async Task<PermissionActionResult> ResetAsync(ManagePermissionsRequest request, string callerUpn, CancellationToken ct)
    {
        if (!TryValidateCommon(request, out var webUri, out var listId, out var commonError))
        {
            return commonError!;
        }

        try
        {
            using var context = await _contextFactory.CreateAsync(webUri, _authProvider);

            var (item, lookupError) = await GetItemAsync(context, listId, request.ItemId!.Value);
            if (item is null)
            {
                return lookupError!;
            }

            if (item.HasUniqueRoleAssignments)
            {
                await item.ResetRoleInheritanceAsync();
                _logger.LogInformation(
                    "RESET durch {Caller}: Vererbung wiederhergestellt auf {Web} Liste {List} Element {Item}.",
                    callerUpn, webUri, listId, request.ItemId);
                return PermissionActionResult.Ok($"Vererbung für Element {request.ItemId} wiederhergestellt.");
            }

            return PermissionActionResult.Ok($"Element {request.ItemId} erbt bereits – keine Änderung nötig.");
        }
        catch (PnP.Core.ServiceException ex)
        {
            return MapServiceException(ex);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unerwarteter Fehler beim Zurücksetzen der Vererbung.");
            return PermissionActionResult.Fail(500, "Unerwarteter Fehler beim Zurücksetzen der Vererbung.");
        }
    }

    private async Task<(IListItem? Item, PermissionActionResult? Error)> GetItemAsync(PnPContext context, Guid listId, int itemId)
    {
        IList list;
        try
        {
            list = await context.Web.Lists.GetByIdAsync(listId, l => l.Id);
        }
        catch (PnP.Core.ServiceException ex) when (StatusOf(ex) == 404)
        {
            return (null, PermissionActionResult.Fail(404, "Liste wurde nicht gefunden."));
        }

        try
        {
            var item = await list.Items.GetByIdAsync(itemId, l => l.Id, l => l.HasUniqueRoleAssignments);
            return (item, null);
        }
        catch (PnP.Core.ServiceException ex) when (StatusOf(ex) == 404)
        {
            return (null, PermissionActionResult.Fail(404, "Listenelement wurde nicht gefunden."));
        }
    }

    private bool TryValidateCommon(ManagePermissionsRequest request, out Uri webUri, out Guid listId, out PermissionActionResult? error)
    {
        webUri = null!;
        listId = Guid.Empty;
        error = null;

        if (string.IsNullOrWhiteSpace(request.WebUrl) ||
            !Uri.TryCreate(request.WebUrl, UriKind.Absolute, out var parsed))
        {
            error = PermissionActionResult.Fail(400, "Feld 'webUrl' fehlt oder ist keine gültige URL.");
            return false;
        }

        if (!IsAllowedWeb(parsed))
        {
            error = PermissionActionResult.Fail(400, "Die angegebene 'webUrl' ist nicht zulässig.");
            return false;
        }

        webUri = parsed;

        if (string.IsNullOrWhiteSpace(request.ListId) || !Guid.TryParse(request.ListId, out listId))
        {
            error = PermissionActionResult.Fail(400, "Feld 'listId' fehlt oder ist keine gültige GUID.");
            return false;
        }

        if (request.ItemId is null or <= 0)
        {
            error = PermissionActionResult.Fail(400, "Feld 'itemId' fehlt oder ist ungültig.");
            return false;
        }

        return true;
    }

    private bool IsAllowedWeb(Uri uri)
    {
        if (uri.Scheme != Uri.UriSchemeHttps)
        {
            return false;
        }

        var allowed = _options.AllowedHostsList;
        if (allowed.Count == 0)
        {
            return uri.Host.EndsWith(".sharepoint.com", StringComparison.OrdinalIgnoreCase);
        }

        return allowed.Contains(uri.Host, StringComparer.OrdinalIgnoreCase);
    }

    private PermissionActionResult MapServiceException(PnP.Core.ServiceException ex)
    {
        var status = StatusOf(ex);
        _logger.LogError(ex, "SharePoint-Dienstfehler (HTTP {Status}).", status);

        return status switch
        {
            404 => PermissionActionResult.Fail(404, "Ressource wurde nicht gefunden."),
            401 or 403 => PermissionActionResult.Fail(502,
                "Zugriff auf SharePoint verweigert – prüfe die Berechtigungen der Managed Identity (Sites.Selected + FullControl)."),
            _ => PermissionActionResult.Fail(500, "SharePoint-Operation fehlgeschlagen."),
        };
    }

    private static int StatusOf(PnP.Core.ServiceException ex) =>
        (ex.Error as PnP.Core.ServiceError)?.HttpResponseCode ?? 0;
}
