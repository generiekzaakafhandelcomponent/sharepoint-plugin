import { TestBed } from '@angular/core/testing';
import { provideHttpClient } from '@angular/common/http';
import { HttpTestingController, provideHttpClientTesting } from '@angular/common/http/testing';
import { ConfigService } from '@valtimo/shared';

import { WorkdocumentsService } from './workdocuments.service';

describe('WorkdocumentsService', () => {
  let service: WorkdocumentsService;

  beforeEach(() => {
    TestBed.configureTestingModule({
      providers: [
        provideHttpClient(),
        provideHttpClientTesting(),
        {
          provide: ConfigService,
          useValue: { config: { valtimoApi: { endpointUri: '/api/' } } },
        },
      ],
    });
    service = TestBed.inject(WorkdocumentsService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });

  it('lists work documents through the case-scoped endpoint', () => {
    const http = TestBed.inject(HttpTestingController);

    service.getWorkDocuments('sharepoint-config', 'case-1', 25, 'cursor').subscribe();

    const request = http.expectOne(
      r => r.url === '/api/v1/plugin/sharepoint/sharepoint-config/documents/case-1/work-documents',
    );
    expect(request.request.params.get('pageSize')).toBe('25');
    expect(request.request.params.get('nextLink')).toBe('cursor');
    http.verify();
  });

  it('uploads a zaak document to SharePoint by its Documenten API configuration and id', () => {
    const http = TestBed.inject(HttpTestingController);

    service
      .uploadFromDocumentenApi('sharepoint-config', 'case-1', {
        fileId: 'doc-1',
        fileName: 'besluit.pdf',
        title: 'Besluit',
        pluginConfigurationId: 'documenten-config',
      })
      .subscribe();

    const request = http.expectOne('/api/v1/plugin/sharepoint/sharepoint-config/documents/case-1/work-documents');
    expect(request.request.method).toBe('POST');
    expect(request.request.body).toEqual({
      documentenApiPluginConfigurationId: 'documenten-config',
      documentenApiDocumentId: 'doc-1',
    });
    http.verify();
  });

  it('stores a SharePoint file in the Documenten API with the given metadata', () => {
    const http = TestBed.inject(HttpTestingController);

    service.storeInDocumentenApi('sharepoint-config', 'case-1', 'item-1', {titel: 'Brief'}).subscribe();

    const request = http.expectOne(
      '/api/v1/plugin/sharepoint/sharepoint-config/documents/case-1/work-documents/item-1/documenten-api',
    );
    expect(request.request.method).toBe('POST');
    expect(request.request.body).toEqual({metadata: {titel: 'Brief'}});
    http.verify();
  });

  it('deletes a SharePoint file of the case', () => {
    const http = TestBed.inject(HttpTestingController);

    service.deleteWorkDocument('sharepoint-config', 'case-1', 'item-1').subscribe();

    const request = http.expectOne('/api/v1/plugin/sharepoint/sharepoint-config/documents/case-1/work-documents/item-1');
    expect(request.request.method).toBe('DELETE');
    http.verify();
  });
});
