# SSL Core Web App – Combined (custom sign-in + SharePoint)

Users sign in **only** with your custom identity. They are never shown a Microsoft sign-in.
The app reaches SharePoint with **its own** identity (an Entra app registration), so no per-user Microsoft token is needed.

```
 Browser ──(1) user + password / customer IdP──►  Auth layer  ──► AppCookie (neutral claims)
                                                      ▲
                                  swap point:  ICredentialAuthenticator  |  OIDC config
 Browser ──(2) pages / downloads──► Controllers ──► ICurrentUser (roles) ──► ISharePointService
                                                                              │ app-only token (client secret / cert)
                                                                              ▼
                                                                    Microsoft Graph ► SharePoint
```

Controllers, views and the SharePoint service only know `ICurrentUser`, `AppRoles` and `ISharePointService`.
Nothing outside `Auth/` knows ASP.NET Identity, OIDC, or any IdP exists.

---

## 0. DMS app inside this portal (new)

Sign in on the portal with **user name + password only** (no Microsoft sign-in). On the home page click **Open DMS**
and the DMS React app (from `DmsSystemAspNetCore`) opens at `/Dms` inside the same session.

```
Browser ──cookie──► /Dms (React app, only for signed-in DmsUser / DmsAdmin)
   │  every call: fetch('/api/sharepoint/_api/...'), '/api/graph/...', '/api/config', '/api/files/open', '/api/photo'
   ▼
Controllers/Dms*.cs ──► DmsApp/DmsProxy ──► app-only token (SharePoint:ClientId + certificate) ──► SharePoint REST / Graph
```

| Piece | Where |
|---|---|
| React source (the original DMS client, 15 small edits) | `ClientApp/` |
| React build output (created by the build, not shipped) | `wwwroot/dms/` |
| Page, config, SharePoint proxy, Graph proxy, licence, file streaming, photos | `Controllers/Dms*.cs`, `DmsApp/` |
| Settings | `appsettings.json` -> `"Dms"` and `"SharePoint"` |

### First run

1. **Node.js 18+** must be installed (only to build the React app once).
2. `dotnet build` - the first build runs `npm install` + `npm run build` in `ClientApp/` and fills `wwwroot/dms/`.
   (Manually: `cd ClientApp && npm install && npm run build`. After changing React code: `dotnet build -p:BuildClient=true`.)
   Run `dotnet build` once **before** `dotnet publish` so the freshly built `wwwroot/dms` is included in the publish output.
3. Do the one-time Entra setup below, then start the **https** profile, sign in, click **Open DMS**.

### One-time Entra / SharePoint setup (required, cannot be skipped)

SharePoint REST (`/_api/...`, which the whole DMS uses) **rejects app-only tokens obtained with a client secret**
("Unsupported app only token"). The app must sign in with a **certificate**:

1. Create a certificate: `tools/New-AppCertificate.ps1 -Password "<strong password>"` (creates `DmsPortal.cer` + `DmsPortal.pfx`).
2. Entra ▸ App registrations ▸ your app (`SharePoint:ClientId`) ▸ Certificates & secrets ▸ **Certificates ▸ Upload `DmsPortal.cer`**.
3. Keep the `.pfx` outside the web root and set (user-secrets or environment variables):
   ```
   dotnet user-secrets set "SharePoint:CertificatePath"     "C:\\secure\\DmsPortal.pfx"
   dotnet user-secrets set "SharePoint:CertificatePassword" "<password>"
   ```
   (When `CertificatePath` is set it is used for both SharePoint and Graph; the client secret is then not needed.)
4. API permissions ▸ **Application** permissions ▸ **Grant admin consent** for all of these:
   * **Office 365 SharePoint Online ▸ `Sites.Selected`** (or `Sites.FullControl.All` for a quick test) - needed for `/_api`
   * **Microsoft Graph ▸ `Sites.Selected`** - site lookup (already in section 2)
   * **Microsoft Graph ▸ `User.Read.All`** - people pickers and profile pictures
   * Microsoft Graph ▸ `Organization.Read.All` - optional, organisation name on the profile page
5. Grant the app a role on **your DMS site** (section 2 below). The DMS creates folders and changes folder permissions, so use
   `"roles": ["fullcontrol"]` (not just `write`).
6. Optional: set `Dms:SiteUrl` (full URL of the DMS site). If empty, the site from `SharePoint:SiteHostname/SitePath/SiteId` is used.

If something is wrong, **Open DMS** shows the reason on screen (for example the certificate message above) instead of a blank page.

### Things that behave differently from the SharePoint web part

* **Who is "the user"?** The portal login. `displayName` / `email` of the portal account are shown in DMS and written into the lists.
  Give each portal user the **same e-mail address as their Microsoft 365 account**, otherwise approvers / assignments written
  in the SharePoint lists will not match them.
* SharePoint's own "Created by / Modified by" show the **app**, not the person (the person is in the DMS columns and the portal log).
* SharePoint no longer enforces per-user permissions: every portal user with role `DmsUser` or `DmsAdmin` can use the whole DMS.
  Restrict who gets those roles. Every write is logged with the real user (`DMS write by ...`).
* Document links (`Open`, approval lists, dashboard) and profile pictures are streamed **through this app**
  (`/api/files/open`, `/api/photo`), so nobody is sent to a Microsoft sign-in page. Only the libraries in `Dms:AllowedLibraries` can be opened.
* **Teams chat sharing** (the "share with a colleague" message in the dashboard) needs a *signed-in Microsoft user*; with an
  app-only identity Microsoft Graph refuses it, so that single action shows an error. Everything else works.
* Fluent UI icons are still loaded from Microsoft's public CDN (no sign-in). Self-host them if users are offline.
* The licence check calls `Dms:License:ApiUrl` from the server. Move it off the free ngrok URL before production.

### What changed in the React code (`ClientApp/src`)

`auth/auth.ts` (cookie instead of MSAL), `spfx-shim/pnp.ts`, `spfx-shim/context.ts` (+`fileOpenUrl`, `photoUrlFor`), `App.tsx`, `vite.config.ts`
(base `/dms/`, output `wwwroot/dms`), removal of `@azure/msal-browser`; `sp.web.currentUser()` replaced by the portal user in
`DocumentList.tsx` / `AnalyticsDashboard.tsx`; SharePoint file and photo URLs replaced by the app's own endpoints in 9 components.

---

## 1. Run it

1. **Rotate the old client secret.** The previous `appsettings.json` contained a live Entra client secret. Treat it as leaked: create a new secret in Entra ▸ App registrations ▸ Certificates & secrets, delete the old one.
2. Store the new one outside source control:
   ```
   dotnet user-secrets set "SharePoint:ClientSecret" "<new secret>"
   ```
   (Production: Azure Key Vault / environment variable `SharePoint__ClientSecret`, or better, a certificate via `SharePoint:CertificatePath`.)
3. Give the app registration access to SharePoint (see §2).
4. Set the site: `SharePoint:SiteId` (best) or `SiteHostname` + `SitePath`.
5. Use the **https** launch profile (the auth cookie is `Secure`). Sign in as `rahul` / `Password123` (admin) or `viewer` / `Password123` (read-only). These are seeded in **Development only**.
6. Open `/SharePoint/Diagnostics` as admin: it prints the resolved site id. If it errors, SharePoint access is the problem, not sign-in.

## 2. Entra app registration (one-time, admin)

1. App registration ▸ **API permissions** ▸ Microsoft Graph ▸ **Application permissions** ▸ add **`Sites.Selected`** ▸ *Grant admin consent*.
   (`Sites.Selected` alone grants access to nothing; that is intended.)
2. An admin then grants the app a role on **your one site** (Graph Explorer with `Sites.FullControl.All`, or PnP `Grant-PnPEntraIDAppSitePermission` (older name: `Grant-PnPAzureADAppSitePermission`)):
   ```
   POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions
   { "roles": ["write"],
     "grantedToIdentities": [ { "application": { "id": "<SharePoint:ClientId>", "displayName": "DMS App" } } ] }
   ```
   Use `"read"` for a read-only app. Folder creation needs `write`.
3. Easiest alternative for a quick test: `Sites.ReadWrite.All` (tenant-wide). The code is safe with either because it only accepts libraries that belong to the configured site.

## 3. Things you need to know before committing to this design

| Consequence | What it means | What the code does about it |
|---|---|---|
| SharePoint sees **one caller: the app** | SharePoint's own per-user permissions, "Created by", and audit log no longer identify the person | `SharePointController` is the authorization boundary (roles `DmsUser` / `DmsAdmin`) and logs every download/create with the real user id |
| Links to SharePoint URLs would trigger a Microsoft sign-in | The old "Open" link in the lists view would have defeated the whole goal | Removed. Files are **streamed through the app** (`/SharePoint/Download`) |
| A user-supplied `driveId` + a broad Graph permission = read any drive in the tenant | Classic IDOR | `EnsureLibraryBelongsToSiteAsync` rejects any library not in the configured site |
| No per-user SharePoint ACLs | If the customer *requires* SharePoint to enforce each user's own permissions, this design cannot do it | That needs a real Entra identity per user (federation/SSO), which is a different architecture |

## 4. Replacing the custom identity with the customer's

The provider is chosen by one setting: `Authentication:Provider`.

**A. Customer has an OpenID Connect IdP (Keycloak, Okta, Auth0, Ping, AD FS 2016+ …)** – configuration only:
```jsonc
"Authentication": {
  "Provider": "Oidc",
  "Oidc": { "Authority": "https://idp.customer.com/realms/x", "ClientId": "dms", "ClientSecret": "<user-secrets>" ,
            "ClaimMapping": { "UserNameClaim": "preferred_username", "RolesClaim": "groups" } },
  "RoleMap": { "DMS-Admins": "DmsAdmin", "DMS-Users": "DmsUser" }
}
```
Register `https://<your-host>/signin-oidc` (redirect) and `https://<your-host>/signout-callback-oidc` (post-logout) at the IdP.
`RoleMap` translates the customer's group names to this app's two roles. When non-empty, unmapped groups are dropped.

**B. Customer has LDAP / a user database / a login API** – write one class, set `Provider` to `Custom`:
```csharp
public sealed class CustomerCredentialAuthenticator : ICredentialAuthenticator
{
    public async Task<CredentialAuthenticationResult> AuthenticateAsync(
        string userName, string password, CancellationToken ct = default)
    {
        // 1. Validate against the customer's system (LDAP bind, REST call, stored procedure ...)
        // 2. On failure:  return CredentialAuthenticationResult.Fail(CredentialFailure.InvalidCredentials);
        // 3. On success:
        return CredentialAuthenticationResult.Success(new AuthenticatedUser(
            SubjectId: "<stable unique id>", UserName: userName, Email: "<email or null>",
            DisplayName: "<display name>", Roles: new[] { AppRoles.DmsUser }));
    }
}
```
Register it in `Program.cs` (the commented line is already there):
`builder.Services.AddScoped<ICredentialAuthenticator, CustomerCredentialAuthenticator>();`

**C. SAML-only customer** – add a SAML2 package, register it with `SignInScheme = AppAuthDefaults.CookieScheme`, and in its ticket-received event call `IExternalClaimsMapper` + `AppPrincipalFactory.Create(...)` exactly as `Auth/Oidc/OidcAuthenticationExtensions.cs` does. (Not included, since it needs a third-party package choice.)

**When the old identity is gone:** delete `Auth/Local/` and the three packages marked in the `.csproj`. Nothing else references them.

> If the customer's IdP *is* Entra ID, users will go through a Microsoft sign-in page (usually silent single sign-on). That is the opposite of this requirement, so it is mentioned only so it isn't a surprise.

Keep per-user data (if you add any later) keyed by `ICurrentUser.UserName`/`Email`, not the provider-specific `Id`: ids change when the provider changes.

## 5. What changed from your sample

**Kept / improved**
- ASP.NET Identity now lives only in `Auth/Local`. It uses `AddIdentityCore` and the app's own cookie, so it behaves the same as every other provider.
- Login: generic error message (no "User Not Found" vs "Invalid Password" leak), **lockout actually counts failures** (`CheckPasswordAsync` alone doesn't), open-redirect-safe `returnUrl`, CSRF tokens on all POSTs, sign-out is POST.
- Password policy raised (10+ chars, upper/lower/digit); the previous policy accepted any 6-character password.
- Secure-by-default: every endpoint needs a signed-in user unless marked `[AllowAnonymous]`.
- One singleton `GraphServiceClient` (the old code built a new credential per request, discarding the token cache). `GraphService` and `SharePointGraphService` were identical, so they are merged into one `ISharePointService`.
- Errors are logged; raw `ex.ToString()` is no longer sent to the browser.
- Added a layout, error page, libraries/files pages, download, and admin-only folder creation (was a state-changing GET).
- `EnsureCreated` and seeding now run in Development only.

**Removed (all were unused or contradicted the goal)**
- Microsoft sign-in flow: `MicrosoftController`, `SsoController`, `InvalidMicrosoftUser` view, `AccountController.MicrosoftLogin/SsoInfo/TestUser`, `ApplicationUser.EntraUpn` (the DB column stays, harmlessly).
- `GraphController` + `GraphService` (referenced a service that was never registered).
- `CertificateAuthController` / `CertificateValidationService`: not wired in, and validation always returned `null`. If mutual-TLS certificate login was the "custom authentication" you meant, tell me; it fits as another `ICredentialAuthenticator`-style provider.
- Unused packages: OpenIddict ×4, Microsoft.Identity.Web, EF InMemory, Certificate auth, CodeGeneration.Design. `Azure.Identity` is now referenced explicitly (it previously arrived through Identity.Web).
- `bin/`, `obj/`, `.vs/` (≈188 MB uncompressed) are not included.

**Existing users:** the earlier sample created users with no roles. Under role-based access they would be signed in but see "no DMS role". In Development the seed fixes `rahul` automatically. Elsewhere add roles, e.g. via a one-off script calling `UserManager.AddToRoleAsync(user, "DmsAdmin")`.

## 6. Known limits

- **Not compiled.** I had no .NET SDK or NuGet access while writing this. Every file passes a C# syntax parse and uses only Graph API shapes your project already compiled against, but please do a first `dotnet build` and expect to fix small things.
- Graph listings read the first page only (up to ~200 items per folder/library). Add paging before using on large libraries.
- Roles are copied into the cookie at sign-in (8 h). A role change takes effect at next sign-in.
- Behind a reverse proxy / load balancer add `UseForwardedHeaders` so HTTPS redirection and the Secure cookie behave.
- Production schema should come from EF migrations (`dotnet ef migrations add Initial`), not `EnsureCreated`.
