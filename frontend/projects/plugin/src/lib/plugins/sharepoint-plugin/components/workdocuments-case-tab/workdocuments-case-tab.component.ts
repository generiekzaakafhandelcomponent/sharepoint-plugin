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

import {ChangeDetectorRef, Component, OnDestroy, OnInit} from "@angular/core";
import {ActivatedRoute} from "@angular/router";
import {PluginManagementService} from "@valtimo/plugin";
import {DocumentenApiDocumentService, DocumentenApiMetadata, DocumentenApiVersionService} from "@valtimo/zgw";
import {TranslateService} from "@ngx-translate/core";
import {BehaviorSubject, combineLatest, map, Observable, of, Subscription} from "rxjs";
import {WorkdocumentsService} from "../../services/workdocuments.service";
import {StoreInDocumentenApiResult, WorkDocument, ZaakDocument} from "../../models";

interface WorkDocumentsContext {
  pluginConfigurationId: string;
}

interface UploadField {
  key: string;
  defaultValue?: string;
  visible?: boolean;
  readonly?: boolean;
}

interface UploadSettings {
  defaultValues: Record<string, string | string[] | undefined>;
  hideFields: string[];
  readonlyFields: string[];
  supportsTrefwoorden: boolean;
}

interface TabMessage {
  type: "success" | "warning" | "error";
  translationKey: string;
  fileName: string;
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
export class WorkdocumentsCaseTabComponent implements OnInit, OnDestroy {
  private readonly documentId: string;
  private readonly subscriptions = new Subscription();
  private context: WorkDocumentsContext | null = null;
  private nextLink: string | undefined;

  workDocuments: WorkDocument[] = [];
  loading = true;
  loadingMore = false;
  error = false;
  hasMore = false;
  sharePointLocation: string | null = null;
  message: TabMessage | null = null;

  pickerOpen = false;
  pickerLoading = false;
  pickerError = false;
  zaakDocuments: ZaakDocument[] = [];
  uploadingFileId: string | null = null;

  storeItem: WorkDocument | null = null;
  readonly storeFile$ = new BehaviorSubject<{name: string} | null>(null);
  readonly storeModalDisabled$ = of(false);
  storing = false;
  uploadSettings: UploadSettings | null = null;

  deleteItem: WorkDocument | null = null;
  deleting = false;

  constructor(
    private readonly route: ActivatedRoute,
    private readonly pluginManagementService: PluginManagementService,
    private readonly workdocumentsService: WorkdocumentsService,
    private readonly translateService: TranslateService,
    private readonly documentenApiDocumentService: DocumentenApiDocumentService,
    private readonly documentenApiVersionService: DocumentenApiVersionService,
    private readonly changeDetectorRef: ChangeDetectorRef,
  ) {
    this.documentId = this.route.snapshot.paramMap.get("documentId") || "";
  }

  ngOnInit(): void {
    this.subscriptions.add(
      this.loadUploadSettings().subscribe({
        next: settings => (this.uploadSettings = settings),
        error: () => (this.uploadSettings = null),
      }),
    );

    this.subscriptions.add(
      this.workdocumentsService.getSharePointLocation(this.documentId).subscribe({
        next: location => (this.sharePointLocation = location),
        error: () => (this.sharePointLocation = null),
      }),
    );

    this.subscriptions.add(
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
      }),
    );
  }

  ngOnDestroy(): void {
    this.subscriptions.unsubscribe();
  }

  loadMore(): void {
    if (!this.hasMore || this.loadingMore) {
      return;
    }
    this.loadingMore = true;
    this.loadPage();
  }

  refresh(): void {
    if (this.loading || this.loadingMore) {
      return;
    }
    this.workDocuments = [];
    this.nextLink = undefined;
    this.hasMore = false;
    this.error = false;
    this.loading = true;
    this.loadPage();
  }

  openPicker(): void {
    this.pickerOpen = true;
    this.pickerLoading = true;
    this.pickerError = false;
    this.subscriptions.add(
      this.workdocumentsService.getZaakDocuments(this.documentId).subscribe({
        next: documents => {
          this.zaakDocuments = documents.filter(document => document.canView !== false);
          this.pickerLoading = false;
        },
        error: () => {
          this.pickerLoading = false;
          this.pickerError = true;
        },
      }),
    );
  }

  closePicker(): void {
    this.pickerOpen = false;
  }

  uploadToSharePoint(zaakDocument: ZaakDocument): void {
    const context = this.context;
    if (!context || this.uploadingFileId) {
      return;
    }
    const fileName = zaakDocument.fileName || zaakDocument.title || "";
    this.uploadingFileId = zaakDocument.fileId;
    this.subscriptions.add(
      this.workdocumentsService
        .uploadFromDocumentenApi(context.pluginConfigurationId, this.documentId, zaakDocument)
        .subscribe({
          next: workDocument => {
            this.uploadingFileId = null;
            this.pickerOpen = false;
            this.message = {type: "success", translationKey: "workDocumentsUploaded", fileName: workDocument.name};
            this.refresh();
          },
          error: () => {
            this.uploadingFileId = null;
            this.pickerOpen = false;
            this.message = {type: "error", translationKey: "workDocumentsUploadError", fileName};
          },
        }),
    );
  }

  openStore(workDocument: WorkDocument): void {
    // The dialog reads the case's upload-field settings when the file arrives, so they load first
    const settings$ = this.uploadSettings ? of(this.uploadSettings) : this.loadUploadSettings();
    this.subscriptions.add(
      settings$.subscribe({
        next: settings => {
          this.uploadSettings = settings;
          // The dialog prefills as soon as the file arrives, so the settings must be bound to it first
          this.changeDetectorRef.detectChanges();
          this.storeItem = workDocument;
          this.storeFile$.next({name: workDocument.name});
        },
        error: () => {
          this.message = {type: "error", translationKey: "workDocumentsStoreError", fileName: workDocument.name};
        },
      }),
    );
  }

  isReadonly(field: string): boolean {
    return !!this.uploadSettings?.readonlyFields.includes(field);
  }

  closeStore(): void {
    if (!this.storing) {
      this.storeItem = null;
    }
  }

  storeInDocumentenApi(metadata: DocumentenApiMetadata): void {
    const context = this.context;
    const workDocument = this.storeItem;
    if (!context || !workDocument || this.storing) {
      return;
    }
    this.storing = true;
    this.subscriptions.add(
      this.workdocumentsService
        .storeInDocumentenApi(context.pluginConfigurationId, this.documentId, workDocument.id, this.toMetadataValues(metadata))
        .subscribe({
          next: result => {
            this.storing = false;
            this.storeItem = null;
            this.message = this.storeResultMessage(result, workDocument.name);
            this.refresh();
          },
          error: () => {
            this.storing = false;
            this.storeItem = null;
            this.message = {type: "error", translationKey: "workDocumentsStoreError", fileName: workDocument.name};
          },
        }),
    );
  }

  confirmDelete(workDocument: WorkDocument): void {
    this.deleteItem = workDocument;
  }

  cancelDelete(): void {
    if (!this.deleting) {
      this.deleteItem = null;
    }
  }

  deleteWorkDocument(): void {
    const context = this.context;
    const workDocument = this.deleteItem;
    if (!context || !workDocument || this.deleting) {
      return;
    }
    this.deleting = true;
    this.subscriptions.add(
      this.workdocumentsService
        .deleteWorkDocument(context.pluginConfigurationId, this.documentId, workDocument.id)
        .subscribe({
          next: () => {
            this.deleting = false;
            this.deleteItem = null;
            this.workDocuments = this.workDocuments.filter(document => document.id !== workDocument.id);
            this.message = {type: "success", translationKey: "workDocumentsDeleted", fileName: workDocument.name};
          },
          error: () => {
            this.deleting = false;
            this.deleteItem = null;
            this.message = {type: "error", translationKey: "workDocumentsDeleteError", fileName: workDocument.name};
          },
        }),
    );
  }

  getFileTypeIcon(name: string): FileTypeIcon {
    const extension = name.split(".").pop()?.toLowerCase() ?? "";
    return FILE_TYPE_ICONS[extension] ?? DEFAULT_FILE_TYPE_ICON;
  }

  formatFileSize(size: number | null): string {
    if (size === null) {
      return "";
    }
    if (size < 1024) {
      return `${size} B`;
    }
    const kilobytes = size / 1024;
    if (kilobytes < 1024) {
      return `${Math.round(kilobytes)} KB`;
    }
    return `${Math.round(kilobytes / 1024)} MB`;
  }

  formatLastModified(dateTime: string | null): string {
    if (!dateTime) {
      return "";
    }
    const date = new Date(dateTime);
    if (isNaN(date.getTime())) {
      return dateTime;
    }
    return new Intl.DateTimeFormat(this.translateService.currentLang, {
      day: "numeric",
      month: "short",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    }).format(date);
  }

  private resolveContext(): Observable<WorkDocumentsContext> {
    return this.pluginManagementService.getPluginConfigurationsByPluginDefinitionKey("sharepoint-plugin").pipe(
      map(configurations => {
        if (configurations.length !== 1 || !configurations[0].id) {
          throw new Error(
            `Expected exactly one SharePoint plugin configuration, found ${configurations.length}. ` +
              "Cannot determine which one to use for this case.",
          );
        }
        return {pluginConfigurationId: configurations[0].id};
      }),
    );
  }

  private loadUploadSettings(): Observable<UploadSettings> {
    const caseDefinitionKey = this.route.snapshot.paramMap.get("caseDefinitionKey");
    return combineLatest([
      this.documentenApiDocumentService.getPrefilledUploadFields(this.documentId),
      caseDefinitionKey ? this.documentenApiVersionService.getSupportedApiFeatures(caseDefinitionKey) : of(null),
    ]).pipe(
      map(([fields, features]) => this.toUploadSettings(fields, !!features?.supportsTrefwoorden)),
    );
  }

  // Same mapping as core's Documenten tab
  private toUploadSettings(fields: UploadField[], supportsTrefwoorden: boolean): UploadSettings {
    const defaultValue = (key: string) => fields.find(field => field.key === key)?.defaultValue;
    return {
      defaultValues: {
        auteur: defaultValue("auteur"),
        vertrouwelijkheidaanduiding: defaultValue("vertrouwelijkheidaanduiding"),
        beschrijving: defaultValue("beschrijving"),
        titel: defaultValue("titel"),
        informatieobjecttype: defaultValue("informatieobjecttype"),
        bestandsnaam: defaultValue("bestandsnaam"),
        taal: defaultValue("taal"),
        status: defaultValue("status"),
        trefwoorden: defaultValue("trefwoorden")
          ?.split(",")
          .map(tag => tag.trim())
          .filter(tag => !!tag),
      },
      hideFields: fields.filter(field => !field.visible).map(field => field.key),
      readonlyFields: fields.filter(field => field.readonly).map(field => field.key),
      supportsTrefwoorden,
    };
  }

  private storeResultMessage(result: StoreInDocumentenApiResult, fileName: string): TabMessage {
    if (!result.stored) {
      return {type: "warning", translationKey: "workDocumentsNotStoredYet", fileName};
    }
    if (!result.removedFromSharePoint) {
      return {type: "warning", translationKey: "workDocumentsStoredNotRemoved", fileName};
    }
    return {type: "success", translationKey: "workDocumentsStored", fileName};
  }

  private toMetadataValues(metadata: DocumentenApiMetadata): Record<string, string> {
    // Same encoding as core's upload: empty values are left out and lists become comma-separated
    return Object.entries(metadata).reduce<Record<string, string>>((values, [key, value]) => {
      if (value !== null && value !== undefined && value !== "") {
        values[key] = Array.isArray(value) ? value.join(",") : String(value);
      }
      return values;
    }, {});
  }

  private loadPage(): void {
    const context = this.context;
    if (!context) {
      return;
    }
    this.subscriptions.add(
      this.workdocumentsService
        .getWorkDocuments(
          context.pluginConfigurationId,
          this.documentId,
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
        }),
    );
  }
}
