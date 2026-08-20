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
import {combineLatest, Observable, of, switchMap} from "rxjs";
import {WorkdocumentsService} from "../../services/workdocuments.service";
import {WorkDocument} from "../../models";

interface WorkDocumentsContext {
  pluginConfigurationId: string;
  docDefinition: string;
  year: string;
  zaaknummer: string;
}

interface FileTypeIcon {
  label: string;
  color: string;
}

const PAGE_SIZE = 25;

const DEFAULT_FILE_TYPE_ICON: FileTypeIcon = {label: "FILE", color: "#9E9E9E"};

const FILE_TYPE_ICONS: Record<string, FileTypeIcon> = {
  pdf: {label: "PDF", color: "#DB4437"},
  doc: {label: "DOC", color: "#2B579A"},
  docx: {label: "DOC", color: "#2B579A"},
  xls: {label: "XLS", color: "#217346"},
  xlsx: {label: "XLS", color: "#217346"},
  csv: {label: "CSV", color: "#217346"},
  ppt: {label: "PPT", color: "#D24726"},
  pptx: {label: "PPT", color: "#D24726"},
  txt: {label: "TXT", color: "#616161"},
  zip: {label: "ZIP", color: "#8A6D3B"},
  rar: {label: "ZIP", color: "#8A6D3B"},
  "7z": {label: "ZIP", color: "#8A6D3B"},
  jpg: {label: "IMG", color: "#7E57C2"},
  jpeg: {label: "IMG", color: "#7E57C2"},
  png: {label: "IMG", color: "#7E57C2"},
  gif: {label: "IMG", color: "#7E57C2"},
  svg: {label: "IMG", color: "#7E57C2"},
};

@Component({
  standalone: false,
  selector: "valtimo-workdocuments-case-tab",
  templateUrl: "./workdocuments-case-tab.component.html",
  styleUrl: "./workdocuments-case-tab.component.css",
})
export class WorkdocumentsCaseTabComponent implements OnInit {
  private readonly documentId: string;
  private context: WorkDocumentsContext | null = null;
  private nextLink: string | undefined;

  workDocuments: WorkDocument[] = [];
  loading = true;
  loadingMore = false;
  error = false;
  hasMore = false;

  constructor(
    private readonly route: ActivatedRoute,
    private readonly documentService: DocumentService,
    private readonly pluginManagementService: PluginManagementService,
    private readonly workdocumentsService: WorkdocumentsService,
  ) {
    this.documentId = this.route.snapshot.paramMap.get("documentId") || "";
  }

  ngOnInit(): void {
    this.resolveContext().subscribe({
      next: context => {
        this.context = context;
        this.loadPage();
      },
      error: error => {
        console.log(error);
        this.loading = false;
        this.error = true;
      },
    });
  }

  loadMore(): void {
    if (!this.hasMore || this.loadingMore) {
      return;
    }
    this.loadingMore = true;
    this.loadPage();
  }

  getFileTypeIcon(name: string): FileTypeIcon {
    const extension = name.split(".").pop()?.toLowerCase() ?? "";
    return FILE_TYPE_ICONS[extension] ?? DEFAULT_FILE_TYPE_ICON;
  }

  private resolveContext(): Observable<WorkDocumentsContext> {
    return combineLatest([
      this.documentService.getDocument(this.documentId),
      this.workdocumentsService.getZaakMetadata(this.documentId),
      this.pluginManagementService.getPluginConfigurationsByPluginDefinitionKey("sharepoint-plugin"),
    ]).pipe(
      switchMap(([document, zaak, configurations]) => {
        const pluginConfigurationId = configurations[0]?.id;
        const startdatum = zaak?.startdatum ? new Date(zaak.startdatum) : null;
        if (!pluginConfigurationId || !document?.definitionId.name || !zaak?.identificatie || !startdatum) {
          throw new Error("Missing SharePoint plugin configuration or zaak properties");
        }
        return of({
          pluginConfigurationId,
          docDefinition: document.definitionId.name,
          year: startdatum.getUTCFullYear().toString(),
          zaaknummer: zaak.identificatie,
        });
      }),
    );
  }

  private loadPage(): void {
    const context = this.context;
    if (!context) {
      return;
    }
    this.workdocumentsService
      .getWorkDocuments(
        context.pluginConfigurationId,
        context.docDefinition,
        context.year,
        context.zaaknummer,
        PAGE_SIZE,
        this.nextLink,
      )
      .subscribe({
        next: page => {
          this.workDocuments = [...this.workDocuments, ...page.documents];
          this.nextLink = page.nextLink ?? undefined;
          this.hasMore = !!page.nextLink;
          this.loading = false;
          this.loadingMore = false;
        },
        error: error => {
          console.log(error);
          this.loading = false;
          this.loadingMore = false;
          this.error = true;
        },
      });
  }
}
