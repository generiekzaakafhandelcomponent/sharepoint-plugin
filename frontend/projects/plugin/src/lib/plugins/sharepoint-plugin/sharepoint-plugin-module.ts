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

import {NgModule} from "@angular/core";
import {CommonModule} from "@angular/common";
import {PluginTranslatePipeModule} from "@valtimo/plugin";
import {FormModule, InputModule as ValtimoInputModule} from "@valtimo/components";
import {SharePointPluginConfigurationComponent} from "./components/sharepoint-plugin-configuration/sharepoint-plugin-configuration.component";
import {CreateZaakFolderConfigurationComponent} from "./components/create-zaak-folder-configuration/create-zaak-folder-configuration.component";
import {WorkdocumentsCaseTabComponent} from "./components/workdocuments-case-tab/workdocuments-case-tab.component";
import {CASE_TAB_TOKEN} from "@valtimo/case";
import {ButtonModule} from "carbon-components-angular";

@NgModule({
  declarations: [
    SharePointPluginConfigurationComponent,
    CreateZaakFolderConfigurationComponent,
    WorkdocumentsCaseTabComponent,
  ],
  imports: [CommonModule, PluginTranslatePipeModule, FormModule, ValtimoInputModule, ButtonModule],
  exports: [
    SharePointPluginConfigurationComponent,
    CreateZaakFolderConfigurationComponent,
    WorkdocumentsCaseTabComponent,
  ],
  providers: [
    {
      provide: CASE_TAB_TOKEN,
      useValue: {
        'werkmap': WorkdocumentsCaseTabComponent,
      },
    }
  ]
})
export class SharePointPluginModule {}
