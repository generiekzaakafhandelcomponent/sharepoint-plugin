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

import {Component, OnInit} from "@angular/core";
import {ActivatedRoute} from "@angular/router";
import {DocumentService} from "@valtimo/document";
import {PluginManagementService} from "@valtimo/plugin";
import {catchError, Observable, of, switchMap, tap} from "rxjs";
import {WorkdocumentsService} from "../../services/workdocuments.service";
import {WorkDocument} from "../../models";
import {ZakenApiZaaktypeLinkService} from "@valtimo/zgw";

interface ZaakDocumentContent {
  zaaktype?: string;
  year?: string;
  zaaknummer?: string;
}

@Component({
  standalone: false,
  selector: "valtimo-workdocuments-case-tab",
  templateUrl: "./workdocuments-case-tab.component.html",
  styleUrl: "./workdocuments-case-tab.component.css",
})
export class WorkdocumentsCaseTabComponent implements OnInit {
  private readonly documentId: string;

  workDocuments$!: Observable<WorkDocument[]>;
  loading = true;
  error = false;

  constructor(
    private readonly route: ActivatedRoute,
    private readonly documentService: DocumentService,
    private readonly zaakTypeLinkService: ZakenApiZaaktypeLinkService,
    private readonly pluginManagementService: PluginManagementService,
    private readonly workdocumentsService: WorkdocumentsService,
  ) {
    this.documentId = this.route.snapshot.paramMap.get("documentId") || "";
  }

  ngOnInit(): void {
    this.workDocuments$ = this.documentService.getDocument(this.documentId).pipe(
      switchMap(document => {
        const content = document.content as ZaakDocumentContent;
        console.log("content");
        console.log(content);
        return this.pluginManagementService
          .getPluginConfigurationsByPluginDefinitionKey("sharepoint-plugin")
          .pipe(
            switchMap(configurations => {
              const pluginConfigurationId = configurations[0]?.id;
              if (!pluginConfigurationId || !content.zaaktype || !content.year || !content.zaaknummer) {
                throw new Error("Missing SharePoint plugin configuration or zaak properties on document");
              }
              return this.workdocumentsService.getWorkDocuments(
                pluginConfigurationId,
                content.zaaktype,
                content.year,
                content.zaaknummer,
              );
            }),
          );
      }),
      tap(() => (this.loading = false)),
      catchError(() => {
        this.loading = false;
        this.error = true;
        return of([]);
      }),
    );
  }
}
