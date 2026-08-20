# Azure AD App Registration Setup

The SharePoint plugin authenticates to Microsoft Graph using the client-credentials flow, backed by an
Azure AD app registration. This page describes how to set that registration up.

## 1. Create the app registration

1. In the [Azure Portal](https://portal.azure.com), go to **Azure Active Directory** → **App registrations**
   → **New registration**.
2. Give it a name (e.g. `sharepoint-plugin`) and register it. No redirect URI is needed, since the plugin
   only uses the client-credentials flow.
3. Note the **Application (client) ID** and **Directory (tenant) ID** on the registration's overview page —
   these map to the plugin's `clientId` and `tenantId` configuration properties.

## 2. Create a client secret

1. Go to **Certificates & secrets** → **Client secrets** → **New client secret**.
2. Copy the secret **value** immediately after creation (it is not shown again) — this maps to the plugin's
   `clientSecret` configuration property.

## 3. Configure API permissions

1. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**.
2. Add one of:
   - **Sites.Read.All** — read access to every SharePoint site in the tenant. Simplest to set up.
   - **Sites.Selected** — no access by default; access is granted per site (see below). Preferred when the
     app registration should only be able to reach specific sites.
3. Click **Grant admin consent for `<tenant>`** and confirm. The permission must show a green checkmark
   under **Status** — without admin consent, the app has no effective access even though the permission is
   listed.

Note: `create-zaak-folder` creates folders in SharePoint, which requires write access
(`Sites.ReadWrite.All`, or a `"write"` role grant when using `Sites.Selected`). Read-only permissions are
only sufficient for listing the work documents case tab.

### If using Sites.Selected

`Sites.Selected` cannot be scoped to a site from the Azure Portal. After granting the permission above,
grant the app registration access to the specific site with a Graph API call (as a user who is a site owner
or Global/SharePoint admin):

```http
POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions
Content-Type: application/json

{
  "roles": ["write"],
  "grantedToIdentities": [
    {
      "application": {
        "id": "<client-id>",
        "displayName": "sharepoint-plugin"
      }
    }
  ]
}
```

Use `"roles": ["read"]` instead if the app registration only needs to list work documents.

## Troubleshooting

If the app registration doesn't have site read permissions consented, Graph returns a "not found" rather
than a "forbidden" for site lookups — a confusing but known behavior. If the plugin fails to resolve the
SharePoint site or drive (`hostname`/`sharePointSiteName`/`baseFolderPath` configuration), verify in the
Azure Portal under your app registration → **API permissions** that `Sites.Read.All` (or `Sites.Selected`)
has been granted **and** that admin consent has been given (green checkmark), before assuming the
site/hostname configuration itself is wrong.
