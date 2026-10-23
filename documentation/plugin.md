# Plugin Documentation

## Overview

The SharePoint plugin integrates Valtimo with Microsoft SharePoint for work document management. It allows a
BPMN process to create a dedicated SharePoint folder for a zaak, and it adds a "Werkdocumenten" case tab that
lists the work documents found in that folder (with thumbnails/file-type icons, locale-aware size and
last-modified formatting, and "load more" pagination through the Microsoft Graph API). From that tab a user can
copy one of the zaak's Documenten API documents into the folder, store a SharePoint document in the Documenten
API, and delete a SharePoint document. A process can delete a SharePoint document with a plugin action, for
example with a file name chosen in a form.

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

### Verwijder document uit SharePoint (`delete-work-document`)

Deletes the file with the given name from the zaak's SharePoint folder, the same folder `create-zaak-folder`
creates. The file name may be a value resolver expression such as `pv:bestandsnaam` or `doc:/bestandsnaam`, so it
can come from a form. A name with a `/` or `\`, or a name that is `.` or `..`, fails the service task. When no file
with that name is in the folder — it was already stored in the Documenten API or deleted from the case tab — the
task completes and a warning is logged.

| Parameter | Type   | Required | Description                                                       |
|-----------|--------|----------|-------------------------------------------------------------------|
| fileName  | string | Yes      | Name of the file in the zaak's SharePoint folder, e.g. `brief.docx`. |

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

The resulting `sharePointFolderLocation` process variable can then be written onto the case document to this value 'sharePointMap' (for
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
- List the work documents in `<document library>/<case definition name>/<year the case was created>/<zaaknummer>`,
  the folder `create-zaak-folder` creates. The backend works out this folder from the case and its zaak, so this
  only works for cases linked to a zaak. A "Load more" button pages through the folder (Microsoft Graph's
  cursor-based paging) and "Refresh" reloads the list.
- **Upload uit Documenten API**: pick one of the zaak's documents from the Documenten API and upload it into the
  folder under its file name. Characters SharePoint does not allow in a file name (`" * : < > ? / \ |`) become
  `_`. When a file with that name already exists, SharePoint keeps both and gives the new file a unique name. The
  document in the Documenten API stays where it is. A document without content is refused.
- **Opslaan in Documenten API**: opens the same metadata dialog as the Documenten tab, with the file name filled in
  and the case's upload field defaults, hidden fields and read-only fields applied.
  On save the file is stored through the case's document upload process, as an upload from the Documenten tab
  would be, so the case needs a document upload process link. Once the Documenten API has stored it, the file is
  removed from SharePoint. When SharePoint cannot remove it, for example because someone has it open, the tab says
  the document was stored and the file stays in SharePoint. When the upload process has started but not stored the
  document yet, the file also stays and the tab says it is not stored yet; check the Documenten tab before storing
  it again, or it is stored twice.
- **Verwijderen**: after confirmation, deletes the file from SharePoint.

#### Permissions

The tab calls these endpoints. Each one checks the user's permissions on the case document, and only acts on
files in the case's own folder:

| Endpoint | Needs |
|----------|-------|
| `GET /api/v1/plugin/sharepoint/{pluginConfigurationId}/documents/{documentId}/work-documents` | `view` on the case |
| `POST /api/v1/plugin/sharepoint/{pluginConfigurationId}/documents/{documentId}/work-documents` | `modify` on the case, and permission to view the Documenten API document |
| `POST /api/v1/plugin/sharepoint/{pluginConfigurationId}/documents/{documentId}/work-documents/{itemId}/documenten-api` | `modify` on the case, and `create` on the new Documenten API document |
| `DELETE /api/v1/plugin/sharepoint/{pluginConfigurationId}/documents/{documentId}/work-documents/{itemId}` | `modify` on the case |

`GET /api/v1/plugin/sharepoint/{pluginConfigurationId}/work-documents`, which takes the case definition, year
and zaaknummer as parameters, is deprecated. It does not check permissions on a case, and the tab no longer uses
it.

### Deleting a SharePoint document from a form

Let a user task form ask for the file name, and delete it in a service task after the user task:

1. In the form, add a text field (or a select with values you configure) with the key `pv.bestandsnaam`.
2. After the user task, add a service task and link it to the `delete-work-document` action:

```json
{
  "activityId": "delete-sharepoint-document",
  "activityType": "bpmn:ServiceTask:start",
  "processLinkType": "plugin",
  "pluginDefinitionKey": "sharepoint-plugin",
  "pluginActionDefinitionKey": "delete-work-document",
  "pluginConfigurationId": "<plugin configuration id>",
  "actionProperties": {
    "fileName": "pv:bestandsnaam"
  }
}
```
