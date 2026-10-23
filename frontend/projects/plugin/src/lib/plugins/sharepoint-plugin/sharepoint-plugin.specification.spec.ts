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

import {sharepointPluginSpecification} from "./sharepoint-plugin.specification";
import {DeleteWorkDocumentConfigurationComponent} from "./components/delete-work-document-configuration/delete-work-document-configuration.component";

describe("sharepointPluginSpecification", () => {
  it("registers the configuration component of the delete-work-document action", () => {
    expect(sharepointPluginSpecification.functionConfigurationComponents?.["delete-work-document"]).toBe(
      DeleteWorkDocumentConfigurationComponent,
    );
  });

  it("names the delete-work-document action in Dutch and English", () => {
    expect(sharepointPluginSpecification.pluginTranslations?.["nl"]?.["delete-work-document"]).toBe(
      "Verwijder document uit SharePoint",
    );
    expect(sharepointPluginSpecification.pluginTranslations?.["en"]?.["delete-work-document"]).toBe(
      "Delete document from SharePoint",
    );
  });
});
