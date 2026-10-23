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

import {Component, Input, NO_ERRORS_SCHEMA, Pipe, PipeTransform} from "@angular/core";
import {ComponentFixture, TestBed} from "@angular/core/testing";
import {By} from "@angular/platform-browser";
import {ActivatedRoute, convertToParamMap} from "@angular/router";
import {PluginManagementService} from "@valtimo/plugin";
import {DocumentenApiDocumentService, DocumentenApiMetadata, DocumentenApiVersionService} from "@valtimo/zgw";
import {TranslateService} from "@ngx-translate/core";
import {Observable, of, Subject, throwError} from "rxjs";
import {WorkdocumentsCaseTabComponent} from "./workdocuments-case-tab.component";
import {WorkdocumentsService} from "../../services/workdocuments.service";
import {WorkDocument, ZaakDocument} from "../../models";

@Pipe({name: "pluginTranslate", standalone: false})
class FakePluginTranslatePipe implements PipeTransform {
  transform(key: string) {
    return of(key);
  }
}

@Component({selector: "valtimo-documenten-api-metadata-modal", template: "", standalone: false})
class StubMetadataModalComponent {
  @Input() defaultValues: Record<string, unknown> = {};
  readonly defaultValuesWhenFileArrived: Record<string, unknown>[] = [];

  @Input() set file$(file$: Observable<{name: string} | null>) {
    file$.subscribe(file => {
      if (file) {
        this.defaultValuesWhenFileArrived.push(this.defaultValues);
      }
    });
  }
}

const brief: WorkDocument = {
  id: "item-1",
  name: "brief.docx",
  webUrl: "https://sharepoint.example/brief.docx",
  size: 2048,
  lastModifiedDateTime: null,
  createdDateTime: null,
  thumbnailUrl: null,
};

const besluit: ZaakDocument = {
  fileId: "doc-1",
  fileName: "besluit.pdf",
  title: "Besluit",
  pluginConfigurationId: "documenten-config",
};

describe("WorkdocumentsCaseTabComponent", () => {
  let fixture: ComponentFixture<WorkdocumentsCaseTabComponent>;
  let component: WorkdocumentsCaseTabComponent;
  let service: jasmine.SpyObj<WorkdocumentsService>;
  let documentenApiDocumentService: jasmine.SpyObj<DocumentenApiDocumentService>;

  beforeEach(() => {
    service = jasmine.createSpyObj<WorkdocumentsService>("WorkdocumentsService", [
      "getWorkDocuments",
      "getSharePointLocation",
      "getZaakDocuments",
      "uploadFromDocumentenApi",
      "storeInDocumentenApi",
      "deleteWorkDocument",
    ]);
    service.getSharePointLocation.and.returnValue(of("https://sharepoint.example/ZAAK-001"));
    service.getWorkDocuments.and.returnValue(of({documents: [brief], nextLink: null}));
    documentenApiDocumentService = jasmine.createSpyObj<DocumentenApiDocumentService>("DocumentenApiDocumentService", [
      "getPrefilledUploadFields",
    ]);
    documentenApiDocumentService.getPrefilledUploadFields.and.returnValue(
      of([
        {key: "vertrouwelijkheidaanduiding", defaultValue: "zaakvertrouwelijk", visible: true, readonly: true},
        {key: "beschrijving", visible: false, readonly: false},
        {key: "trefwoorden", defaultValue: "concept, brief", visible: true, readonly: false},
      ] as any),
    );

    TestBed.configureTestingModule({
      declarations: [WorkdocumentsCaseTabComponent, FakePluginTranslatePipe, StubMetadataModalComponent],
      providers: [
        {
          provide: ActivatedRoute,
          useValue: {snapshot: {paramMap: convertToParamMap({documentId: "case-1", caseDefinitionKey: "example"})}},
        },
        {provide: DocumentenApiDocumentService, useValue: documentenApiDocumentService},
        {
          provide: DocumentenApiVersionService,
          useValue: {getSupportedApiFeatures: (key: string) => of({supportsTrefwoorden: key === "example"})},
        },
        {
          provide: PluginManagementService,
          useValue: {getPluginConfigurationsByPluginDefinitionKey: () => of([{id: "sharepoint-config"}])},
        },
        {provide: WorkdocumentsService, useValue: service},
        {provide: TranslateService, useValue: {currentLang: "nl"}},
      ],
      schemas: [NO_ERRORS_SCHEMA],
    });
    fixture = TestBed.createComponent(WorkdocumentsCaseTabComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  const click = (selector: string) => {
    fixture.debugElement.query(By.css(selector)).nativeElement.click();
    fixture.detectChanges();
  };

  const messageText = () => fixture.debugElement.query(By.css(".workdocuments-message"))?.nativeElement.textContent.trim();

  it("lists the case's work documents through the case-scoped endpoint", () => {
    expect(service.getWorkDocuments).toHaveBeenCalledWith("sharepoint-config", "case-1", 25, undefined);
    const link = fixture.debugElement.query(By.css(".workdocuments-table a")).nativeElement;
    expect(link.textContent).toContain("brief.docx");
    expect(link.getAttribute("href")).toBe("https://sharepoint.example/brief.docx");
  });

  it("uploads a picked Documenten API document of the zaak to SharePoint and refreshes the list", () => {
    service.getZaakDocuments.and.returnValue(of([besluit]));
    service.uploadFromDocumentenApi.and.returnValue(of({...brief, id: "item-2", name: "besluit.pdf"}));

    click(".workdocuments-open-picker");
    expect(service.getZaakDocuments).toHaveBeenCalledWith("case-1");
    click(".workdocuments-upload");

    expect(service.uploadFromDocumentenApi).toHaveBeenCalledWith("sharepoint-config", "case-1", besluit);
    expect(service.getWorkDocuments).toHaveBeenCalledTimes(2);
    expect(component.pickerOpen).toBeFalse();
    expect(messageText()).toBe("workDocumentsUploaded besluit.pdf");
  });

  it("offers only the zaak documents the user may view", () => {
    service.getZaakDocuments.and.returnValue(of([besluit, {...besluit, fileId: "doc-2", canView: false}]));

    click(".workdocuments-open-picker");

    expect(component.zaakDocuments).toEqual([besluit]);
  });

  it("names the document when uploading to SharePoint fails", () => {
    service.getZaakDocuments.and.returnValue(of([besluit]));
    service.uploadFromDocumentenApi.and.returnValue(throwError(() => new Error("409")));

    click(".workdocuments-open-picker");
    click(".workdocuments-upload");

    expect(messageText()).toBe("workDocumentsUploadError besluit.pdf");
  });

  it("opens the Documenten API metadata dialog for a SharePoint file with its name filled in", () => {
    const modal = () => fixture.debugElement.query(By.css("valtimo-documenten-api-metadata-modal"));
    expect(modal().properties["open"]).toBeFalse();

    click(".workdocuments-store");

    expect(modal().properties["open"]).toBeTrue();
    let file: {name: string} | null = null;
    component.storeFile$.subscribe(value => (file = value));
    expect(file!).toEqual({name: "brief.docx"});
  });

  it("gives the metadata dialog the case's upload-field defaults, hidden and read-only fields", () => {
    click(".workdocuments-store");

    const modal = fixture.debugElement.query(By.css("valtimo-documenten-api-metadata-modal"));
    expect(documentenApiDocumentService.getPrefilledUploadFields).toHaveBeenCalledWith("case-1");
    const defaultValues = (modal.componentInstance as StubMetadataModalComponent).defaultValues;
    expect(defaultValues["vertrouwelijkheidaanduiding"]).toBe("zaakvertrouwelijk");
    expect(defaultValues["trefwoorden"]).toEqual(["concept", "brief"]);
    expect(modal.properties["disableConfidentialityLevel"]).toBeTrue();
    expect(modal.properties["disableDescription"]).toBeFalse();
    expect(modal.properties["hideFields"]).toEqual(["beschrijving"]);
    expect(modal.properties["supportsTrefwoorden"]).toBeTrue();
  });

  it("binds the upload-field defaults before the dialog prefills, also when they arrive after the tab opened", () => {
    const uploadFields$ = new Subject<any>();
    documentenApiDocumentService.getPrefilledUploadFields.and.returnValue(uploadFields$);
    fixture = TestBed.createComponent(WorkdocumentsCaseTabComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();

    click(".workdocuments-store");
    uploadFields$.next([{key: "vertrouwelijkheidaanduiding", defaultValue: "zaakvertrouwelijk", visible: true, readonly: true}]);

    const modal: StubMetadataModalComponent = fixture.debugElement.query(
      By.directive(StubMetadataModalComponent),
    ).componentInstance;
    expect(modal.defaultValuesWhenFileArrived.length).toBe(1);
    expect(modal.defaultValuesWhenFileArrived[0]["vertrouwelijkheidaanduiding"]).toBe("zaakvertrouwelijk");
  });

  it("does not open the metadata dialog when the case's upload-field settings cannot be loaded", () => {
    documentenApiDocumentService.getPrefilledUploadFields.and.returnValue(throwError(() => new Error("500")));
    fixture = TestBed.createComponent(WorkdocumentsCaseTabComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();

    click(".workdocuments-store");

    expect(component.storeItem).toBeNull();
    expect(messageText()).toBe("workDocumentsStoreError brief.docx");
  });

  it("stores a SharePoint file with the dialog's metadata and reports it removed from SharePoint", () => {
    service.storeInDocumentenApi.and.returnValue(of({stored: true, removedFromSharePoint: true}));
    click(".workdocuments-store");

    component.storeInDocumentenApi({
      titel: "Brief",
      auteur: "jan@example.com",
      taal: "nld",
      informatieobjecttype: "https://catalogi.example/type",
      creatiedatum: "2026-10-23",
      trefwoorden: ["concept", "brief"],
      beschrijving: "",
      documentUrlProcessVariable: null,
    } as unknown as DocumentenApiMetadata);
    fixture.detectChanges();

    expect(service.storeInDocumentenApi).toHaveBeenCalledWith("sharepoint-config", "case-1", "item-1", {
      titel: "Brief",
      auteur: "jan@example.com",
      taal: "nld",
      informatieobjecttype: "https://catalogi.example/type",
      creatiedatum: "2026-10-23",
      trefwoorden: "concept,brief",
    });
    expect(component.storeItem).toBeNull();
    expect(service.getWorkDocuments).toHaveBeenCalledTimes(2);
    expect(messageText()).toBe("workDocumentsStored brief.docx");
  });

  it("warns when a stored file could not be removed from SharePoint", () => {
    service.storeInDocumentenApi.and.returnValue(of({stored: true, removedFromSharePoint: false}));
    click(".workdocuments-store");

    component.storeInDocumentenApi({titel: "Brief"} as DocumentenApiMetadata);
    fixture.detectChanges();

    expect(messageText()).toBe("workDocumentsStoredNotRemoved brief.docx");
    expect(fixture.debugElement.query(By.css(".workdocuments-message-warning"))).toBeTruthy();
  });

  it("says a file is not stored yet without calling it a failure", () => {
    service.storeInDocumentenApi.and.returnValue(of({stored: false, removedFromSharePoint: false}));
    click(".workdocuments-store");

    component.storeInDocumentenApi({titel: "Brief"} as DocumentenApiMetadata);
    fixture.detectChanges();

    expect(messageText()).toBe("workDocumentsNotStoredYet brief.docx");
    expect(fixture.debugElement.query(By.css(".workdocuments-message-error"))).toBeNull();
  });

  it("shows an error and keeps the file when storing in the Documenten API fails", () => {
    service.storeInDocumentenApi.and.returnValue(throwError(() => new Error("409")));
    click(".workdocuments-store");

    component.storeInDocumentenApi({titel: "Brief"} as DocumentenApiMetadata);
    fixture.detectChanges();

    expect(messageText()).toBe("workDocumentsStoreError brief.docx");
    expect(component.workDocuments).toEqual([brief]);
  });

  it("deletes a SharePoint file only after confirming, and removes it from the list", () => {
    service.deleteWorkDocument.and.returnValue(of(undefined));

    click(".workdocuments-delete");
    expect(service.deleteWorkDocument).not.toHaveBeenCalled();
    click(".workdocuments-delete-confirm");

    expect(service.deleteWorkDocument).toHaveBeenCalledWith("sharepoint-config", "case-1", "item-1");
    expect(component.workDocuments).toEqual([]);
    expect(messageText()).toBe("workDocumentsDeleted brief.docx");
  });

  it("keeps the file listed and names it when SharePoint refuses the delete", () => {
    service.deleteWorkDocument.and.returnValue(throwError(() => new Error("423")));

    click(".workdocuments-delete");
    click(".workdocuments-delete-confirm");

    expect(component.workDocuments).toEqual([brief]);
    expect(messageText()).toBe("workDocumentsDeleteError brief.docx");
  });
});
