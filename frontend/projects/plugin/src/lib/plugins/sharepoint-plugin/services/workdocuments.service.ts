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

import {Injectable} from "@angular/core";
import {HttpClient} from "@angular/common/http";
import {ConfigService} from "@valtimo/shared";
import {Observable} from "rxjs";
import {WorkDocumentPage, ZaakMetadata} from "../models";

@Injectable({
  providedIn: "root",
})
export class WorkdocumentsService {
  private readonly valtimoEndpointUri: string;

  constructor(
    private readonly http: HttpClient,
    private readonly configService: ConfigService,
  ) {
    this.valtimoEndpointUri = this.configService.config.valtimoApi.endpointUri;
  }

  getWorkDocuments(
    pluginConfigurationId: string,
    docDefinition: string,
    year: string,
    zaaknummer: string,
    pageSize?: number,
    nextLink?: string,
  ): Observable<WorkDocumentPage> {
    const params: Record<string, string> = {docDefinition, year, zaaknummer};
    if (pageSize) {
      params["pageSize"] = pageSize.toString();
    }
    if (nextLink) {
      params["nextLink"] = nextLink;
    }
    return this.http.get<WorkDocumentPage>(
      `${this.valtimoEndpointUri}v1/plugin/sharepoint/${pluginConfigurationId}/work-documents`,
      {params},
    );
  }

  getZaakMetadata(documentId: string): Observable<ZaakMetadata> {
    return this.http.get<ZaakMetadata>(
      `${this.valtimoEndpointUri}v1/zaken-api/document/${documentId}/zaak`,
    );
  }

  getSharePointLocation(documentId: string): Observable<string> {
    return this.http.get(
      `${this.valtimoEndpointUri}v1/plugin/sharepoint/documents/${documentId}/sharepoint-location`,
      {responseType: "text"}
    );
  }
}
