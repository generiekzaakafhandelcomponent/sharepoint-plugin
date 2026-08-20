/*
 * Copyright 2026 Ritense BV, the Netherlands.
 *
 * Licensed under EUPL, Version 1.2 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * https://joinup.ec.europa.eu/collection/eupl/eupl-text-eupl-12
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" basis,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import {PluginSpecification} from "@valtimo/plugin";
import {SharePointPluginConfigurationComponent} from "./components/sharepoint-plugin-configuration/sharepoint-plugin-configuration.component";
import {SHAREPOINT_PLUGIN_LOGO_BASE64} from "./assets";
import {CreateZaakFolderConfigurationComponent} from "./components/create-zaak-folder-configuration/create-zaak-folder-configuration.component";
import {ListWorkDocumentsConfigurationComponent} from "./components/list-work-documents-configuration/list-work-documents-configuration.component";

const sharepointPluginSpecification: PluginSpecification = {
  pluginId: "sharepoint-plugin",
  pluginConfigurationComponent: SharePointPluginConfigurationComponent,
  pluginLogoBase64: SHAREPOINT_PLUGIN_LOGO_BASE64,
  functionConfigurationComponents: {
    "create-zaak-folder": CreateZaakFolderConfigurationComponent,
    "list-work-documents": ListWorkDocumentsConfigurationComponent,
  },
  pluginTranslations: {
    nl: {
      title: "SharePoint Plugin",
      description: "Plugin voor integratie met Microsoft SharePoint voor werkdocumentenbeheer.",
      configurationTitle: "Configuratienaam",
      tenantId: "Tenant ID",
      clientId: "Client ID",
      clientSecret: "Client Secret",
      hostname: "Hostname",
      sharePointSiteName: "SharePoint Site Name",
      baseFolderPath: "Basismap",
      "create-zaak-folder": "Maak zaakmap aan in SharePoint",
      createZaakFolderDescription:
        "Maakt automatisch een map aan in SharePoint voor de werkdocumenten van de zaak.",
      dossierDefinitionName: "Dossier definitie naam",
      yearVariable: "Jaar waarin dossier is aangemaakt",
      zaaknummerVariable: "Zaaknummer",
      "list-work-documents": "Haal werkdocumenten op uit SharePoint",
      listWorkDocumentsDescription:
        "Haalt alle werkdocumenten op uit SharePoint voor de zaak en slaat ze op als procesvariabele.",
      sharePointZaakFolderProcessVariable: "Process variable met sharepointlocatie",
      resultVariable: "Process variabele met het resultaat",
    },
    en: {
      title: "SharePoint Plugin",
      description: "Plugin for integration with Microsoft SharePoint for work document management.",
      configurationTitle: "Configuration Name",
      tenantId: "Tenant ID",
      clientId: "Client ID",
      clientSecret: "Client Secret",
      hostname: "Hostname",
      sharePointSiteName: "SharePoint Site Name",
      baseFolderPath: "Base Folder Path",
      "create-zaak-folder": "Create case folder in SharePoint",
      createZaakFolderDescription:
        "Automatically creates a folder in SharePoint for the case's work documents.",
      dossierDefinitionName: "Case type",
      yearVariable: "Year case created",
      zaaknummerVariable: "Case number",
      "list-work-documents": "Retrieve work documents from SharePoint",
      listWorkDocumentsDescription:
        "Retrieves all work documents from SharePoint for the case and stores them as a process variable.",
      sharePointZaakFolderProcessVariable: "Process variable with sharepointlocation",
      resultVariable: "Process variabele with the result",
    },
  },
};

export {sharepointPluginSpecification};
