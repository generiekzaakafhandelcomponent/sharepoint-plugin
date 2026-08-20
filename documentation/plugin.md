# Plugin Documentation

## Overview

The SharePoint plugin integrates Valtimo with Microsoft SharePoint for work document management. It allows a
BPMN process to create a dedicated SharePoint folder for a zaak, and it adds a "Werkdocumenten" case tab that
lists the work documents found in that folder (with thumbnails/file-type icons, locale-aware size and
last-modified formatting, and "load more" pagination through the Microsoft Graph API).

Under the hood it uses [Microsoft Graph](https://learn.microsoft.com/en-us/graph/) (via the
`microsoft-graph` Java SDK) to talk to SharePoint, authenticating with an Azure AD app registration through
client-credentials (`azure-identity`).

## Dependencies

### Backend

```kotlin
dependencies {
    implementation("com.ritense.valtimoplugins:sharepoint-plugin:0.0.1")
}
```

### Frontend

```json
{
  "dependencies": {
    "@valtimo-plugins/sharepoint-plugin": "0.0.1"
  }
}
```

In your `app.module.ts`, register the plugin specification and, to show the work documents case tab, add
`WorkdocumentsCaseTabComponent` to the case tab factory under the content key used in your case tab
configuration (`sharepoint-workdocuments` in the example case):

```typescript
import {
    SharePointPluginModule,
    sharepointPluginSpecification,
    WorkdocumentsCaseTabComponent,
} from '@valtimo-plugins/sharepoint-plugin';
import {CaseModule, CaseDetailTabSummaryComponent, DefaultTabs} from '@valtimo/case';
import {PLUGINS_TOKEN} from '@valtimo/plugin';

export function tabsFactory() {
  return new Map<string, object>([
    [DefaultTabs.summary, CaseDetailTabSummaryComponent],
    ['sharepoint-workdocuments', WorkdocumentsCaseTabComponent],
  ]);
}

@NgModule({
    imports: [
        SharePointPluginModule,
        CaseModule.forRoot(tabsFactory),
    ],
    providers: [
        {
            provide: PLUGINS_TOKEN,
            useValue: [
                sharepointPluginSpecification,
            ]
        }
    ]
})
```

## Configuration

The plugin configuration holds the Azure AD app registration credentials and the SharePoint site/library to
use.

| Property           | Type   | Required | Description                                                                          |
|---------------------|--------|----------|---------------------------------------------------------------------------------------|
| tenantId            | string | Yes      | Azure AD tenant ID of the app registration used to authenticate against Graph.        |
| clientId            | string | Yes      | Client (application) ID of the Azure AD app registration.                             |
| clientSecret        | string | Yes      | Client secret of the Azure AD app registration. Stored encrypted.                     |
| hostname            | string | Yes      | SharePoint hostname, e.g. `contoso.sharepoint.com`.                                    |
| sharePointSiteName  | string | Yes      | Path of the SharePoint site under `/sites/`, e.g. `mysite`.                            |
| baseFolderPath      | string | Yes      | Name of the SharePoint document library (drive) in the site to store documents in.     |

## Actions

### Maak zaakmap aan in SharePoint (`create-zaak-folder`)

Creates a folder in SharePoint for the work documents of a zaak, and stores the created folder's SharePoint
URL in a process variable. The folder is created at
`<document library>/<case definition name>/<year the case was created>/<zaaknummer>`; if the folder already
exists, the existing folder's location is returned instead of failing.

The zaaknummer is resolved automatically from the zaak linked to the process (via the `identificatie` zaak
value resolver), and the year is taken from the case's creation date, so neither needs to be supplied as an
action property.

| Parameter                          | Type   | Required | Description                                                                 |
|-------------------------------------|--------|----------|-------------------------------------------------------------------------------|
| sharePointZaakFolderProcessVariable | string | Yes      | Name of the process variable that the created folder's SharePoint URL is stored in. |

## Usage

### Creating the zaak folder from a process

Add a service task to the BPMN process and link it to the `create-zaak-folder` plugin action, for example:

```json
{
  "activityId": "create-zaak-folder",
  "activityType": "bpmn:ServiceTask:start",
  "processLinkType": "plugin",
  "pluginDefinitionKey": "sharepoint-plugin",
  "pluginActionDefinitionKey": "create-zaak-folder",
  "pluginConfigurationId": "<plugin configuration id>",
  "actionProperties": {
    "sharePointZaakFolderProcessVariable": "sharePointFolderLocation"
  }
}
```

The resulting `sharePointFolderLocation` process variable can then be written onto the case document (for
example via a `camunda:executionListener` calling `valueResolverDelegateService.handleValue(execution,
"doc:/sharePointMap", sharePointFolderLocation)`), so the case tab described below can show a link to it.

### Showing work documents on the case

Once `WorkdocumentsCaseTabComponent` is registered (see [Dependencies](#dependencies)) and a `"custom"` case
tab is configured with `contentKey` `sharepoint-workdocuments`, e.g.:

```json
{
  "name": "Werkdocumenten",
  "key": "workdocuments",
  "type": "custom",
  "contentKey": "sharepoint-workdocuments"
}
```

the tab will:

- Show a link to the SharePoint folder, if `doc:/sharePointMap` is set on the case document.
- List the work documents in `<document library>/<case definition name>/<zaak start year>/<zaaknummer>`,
  resolving the zaak's `identificatie` and `startdatum` through the zaak metadata endpoint
  (`GET /api/v1/zaken-api/document/{documentId}/zaak`), so this only works for cases linked to a zaak. Note
  that this is the zaak's start year, whereas `create-zaak-folder` uses the case document's creation year —
  the folder only lines up between the two actions when those years match.
- Fetch documents from `GET /api/v1/plugin/sharepoint/{pluginConfigurationId}/work-documents`, paginating via
  a "Load more" button (Microsoft Graph's cursor-based paging), and offering a "Refresh" button to reload the
  list from scratch.
