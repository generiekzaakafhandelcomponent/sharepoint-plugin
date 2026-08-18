import { TestBed } from '@angular/core/testing';
import { provideHttpClient } from '@angular/common/http';
import { provideHttpClientTesting } from '@angular/common/http/testing';
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
});
