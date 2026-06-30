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

import {AccessControlManagementModule} from "@valtimo/access-control-management";
import {AccountModule} from "@valtimo/account";
import {AnalyseModule} from "@valtimo/analyse";
import {AppComponent} from "./app.component";
import {AppRoutingModule} from "./app-routing.module";
import {
  BesluitenApiPluginModule,
  besluitenApiPluginSpecification,
  CatalogiApiPluginModule,
  catalogiApiPluginSpecification,
  DocumentenApiPluginModule,
  documentenApiPluginSpecification,
  NotificatiesApiPluginModule,
  notificatiesApiPluginSpecification,
  ObjectenApiPluginModule,
  objectenApiPluginSpecification,
  ObjectTokenAuthenticationPluginModule,
  objectTokenAuthenticationPluginSpecification,
  ObjecttypenApiPluginModule,
  objecttypenApiPluginSpecification,
  OpenNotificatiesPluginModule,
  openNotificatiesPluginSpecification,
  OpenZaakPluginModule,
  openZaakPluginSpecification,
  PLUGINS_TOKEN,
  ZakenApiPluginModule,
  zakenApiPluginSpecification,
} from "@valtimo/plugin";
import {BootstrapModule} from "@valtimo/bootstrap";
import {
  BpmnJsDiagramModule,
  enableCustomFormioComponents,
  MenuModule,
  registerFormioFileSelectorComponent,
  registerFormioUploadComponent,
  registerFormioValueResolverSelectorComponent,
  ValuePathSelectorComponent,
  WidgetModule,
} from "@valtimo/components";
import {BrowserModule} from "@angular/platform-browser";
import {BuildingBlockManagementModule} from "@valtimo/building-block-management";
import {CaseManagementModule} from "@valtimo/case-management";
import {CaseMigrationModule} from "@valtimo/case-migration";
import {ChoiceFieldModule} from "@valtimo/choice-field";
import {CommonModule} from "@angular/common";
import {ConfigModule, ConfigService, CustomMultiTranslateHttpLoaderFactory, LocalizationService} from "@valtimo/shared";
import {DashboardManagementModule} from "@valtimo/dashboard-management";
import {DashboardModule} from "@valtimo/dashboard";
import {DecisionModule} from "@valtimo/decision";
import {
  CaseDetailTabAuditComponent,
  CaseDetailTabDocumentsComponent,
  CaseDetailTabProgressComponent,
  CaseDetailTabSummaryComponent,
  CaseModule,
  DefaultTabs,
} from "@valtimo/case";
import {DocumentModule} from "@valtimo/document";
import {FormManagementModule} from "@valtimo/form-management";
import {FormModule} from "@valtimo/form";
import {FormsModule, ReactiveFormsModule} from "@angular/forms";
import {HttpBackend, HttpClient, HttpClientModule} from "@angular/common/http";
import {Injector, NgModule} from "@angular/core";
import {LayoutModule, TranslationManagementModule} from "@valtimo/layout";
import {LoggerModule} from "ngx-logger";
import {LoggingModule} from "@valtimo/logging";
import {MigrationModule} from "@valtimo/migration";
import {MilestoneModule} from "@valtimo/milestone";
import {ObjectManagementModule} from "@valtimo/object-management";
import {ObjectModule} from "@valtimo/object";
import {PluginManagementModule} from "@valtimo/plugin-management";
import {ProcessLinkModule} from "@valtimo/process-link";
import {ProcessManagementModule} from "@valtimo/process-management";
import {ProcessModule} from "@valtimo/process";
import {ResourceModule} from "@valtimo/resource";
import {SecurityModule} from "@valtimo/security";
import {SseModule} from "@valtimo/sse";
import {SwaggerModule} from "@valtimo/swagger";
import {TaskModule} from "@valtimo/task";
import {TeamsModule} from "@valtimo/teams";
import {TranslateLoader, TranslateModule} from "@ngx-translate/core";
import {environment} from "../environments/environment";
import {registerDocumentenApiFormioUploadComponent, ZgwModule} from "@valtimo/zgw";

import {SamplePluginModule, samplePluginSpecification,} from "@valtimo-plugins/sample-plugin";

export function tabsFactory() {
  return new Map<string, object>([
    [DefaultTabs.summary, CaseDetailTabSummaryComponent],
    [DefaultTabs.progress, CaseDetailTabProgressComponent],
    [DefaultTabs.audit, CaseDetailTabAuditComponent],
    [DefaultTabs.documents, CaseDetailTabDocumentsComponent],
  ]);
}

@NgModule({
  declarations: [AppComponent],
  imports: [
    environment.authentication.module,
    AccessControlManagementModule,
    AccountModule,
    AnalyseModule,
    AppRoutingModule,
    BootstrapModule,
    BpmnJsDiagramModule,
    BrowserModule,
    BuildingBlockManagementModule,
    CaseManagementModule,
    CaseMigrationModule,
    CaseModule.forRoot(tabsFactory),
    BesluitenApiPluginModule,
    CatalogiApiPluginModule,
    ChoiceFieldModule,
    CommonModule,
    ConfigModule.forRoot(environment),
    DashboardManagementModule,
    DashboardModule,
    DecisionModule,
    DocumentenApiPluginModule,
    DocumentModule,
    FormManagementModule,
    FormModule,
    FormsModule,
    HttpClientModule,
    LayoutModule,
    LoggerModule.forRoot(environment.logger),
    LoggingModule,
    MenuModule,
    MigrationModule,
    MilestoneModule,
    NotificatiesApiPluginModule,
    ObjectenApiPluginModule,
    ObjectManagementModule,
    ObjectModule,
    ObjectTokenAuthenticationPluginModule,
    ObjecttypenApiPluginModule,
    OpenNotificatiesPluginModule,
    OpenZaakPluginModule,
    PluginManagementModule,
    ProcessLinkModule,
    ProcessManagementModule,
    ProcessModule,
    ReactiveFormsModule,
    ResourceModule,
    SamplePluginModule,
    SecurityModule,
    SseModule,
    SwaggerModule,
    TaskModule,
    TeamsModule,
    TranslateModule.forRoot({
      loader: {
        provide: TranslateLoader,
        useFactory: CustomMultiTranslateHttpLoaderFactory,
        deps: [HttpBackend, HttpClient, ConfigService, LocalizationService],
      },
    }),
    TranslationManagementModule,
    ValuePathSelectorComponent,
    WidgetModule,
    ZakenApiPluginModule,
    ZgwModule,
  ],
  providers: [
    {
      provide: PLUGINS_TOKEN,
      useValue: [
        besluitenApiPluginSpecification,
        catalogiApiPluginSpecification,
        documentenApiPluginSpecification,
        notificatiesApiPluginSpecification,
        objectTokenAuthenticationPluginSpecification,
        objectenApiPluginSpecification,
        objecttypenApiPluginSpecification,
        openNotificatiesPluginSpecification,
        openZaakPluginSpecification,
        samplePluginSpecification,
        zakenApiPluginSpecification,
      ],
    },
  ],
  bootstrap: [AppComponent],
})
export class AppModule {
  constructor(injector: Injector) {
    enableCustomFormioComponents(injector);
    registerFormioUploadComponent(injector);
    registerFormioFileSelectorComponent(injector);
    registerDocumentenApiFormioUploadComponent(injector);
    registerFormioValueResolverSelectorComponent(injector);
  }
}
