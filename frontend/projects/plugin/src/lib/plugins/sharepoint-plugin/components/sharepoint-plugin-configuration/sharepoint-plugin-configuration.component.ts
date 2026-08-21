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

import {Component, EventEmitter, Input, OnDestroy, OnInit, Output} from "@angular/core";
import {HttpClient} from "@angular/common/http";
import {ConfigService} from "@valtimo/shared";
import {PluginConfigurationComponent, PluginConfigurationData} from "@valtimo/plugin";
import {BehaviorSubject, combineLatest, Observable, Subscription, switchMap, take} from "rxjs";
import {SharePointPluginConfig, TestConnectionResult} from "../../models";

@Component({
  standalone: false,
  selector: "valtimo-sharepoint-plugin-configuration",
  templateUrl: "./sharepoint-plugin-configuration.component.html",
})
export class SharePointPluginConfigurationComponent
  implements PluginConfigurationComponent, OnInit, OnDestroy
{
  @Input() save$!: Observable<void>;
  @Input() disabled$!: Observable<boolean>;
  @Input() pluginId!: string;
  @Input() prefillConfiguration$!: Observable<SharePointPluginConfig>;
  @Output() valid: EventEmitter<boolean> = new EventEmitter<boolean>();
  @Output() configuration: EventEmitter<PluginConfigurationData> =
    new EventEmitter<PluginConfigurationData>();

  private saveSubscription!: Subscription;
  private readonly valtimoEndpointUri: string;
  private readonly formValue$ = new BehaviorSubject<SharePointPluginConfig | null>(null);
  private readonly valid$ = new BehaviorSubject<boolean>(false);

  testingConnection = false;
  testConnectionResult: TestConnectionResult | null = null;

  constructor(
    private readonly http: HttpClient,
    private readonly configService: ConfigService,
  ) {
    this.valtimoEndpointUri = this.configService.config.valtimoApi.endpointUri;
  }

  ngOnInit(): void {
    this.openSaveSubscription();
  }

  ngOnDestroy(): void {
    this.saveSubscription?.unsubscribe();
  }

  formValueChange(formValue: SharePointPluginConfig): void {
    this.formValue$.next(formValue);
    this.testConnectionResult = null;
    this.handleValid(formValue);
  }

  canTestConnection(): boolean {
    const formValue = this.formValue$.value;
    return !!(
      formValue?.tenantId &&
      formValue?.clientId &&
      formValue?.clientSecret &&
      formValue?.hostname &&
      formValue?.sharePointSiteName &&
      formValue?.baseFolderPath
    );
  }

  testConnection(): void {
    const formValue = this.formValue$.value;
    if (!formValue || this.testingConnection) {
      return;
    }
    this.testingConnection = true;
    this.testConnectionResult = null;
    this.http
      .post<TestConnectionResult>(`${this.valtimoEndpointUri}v1/plugin/sharepoint/test-connection`, {
        tenantId: formValue.tenantId,
        clientId: formValue.clientId,
        clientSecret: formValue.clientSecret,
        hostname: formValue.hostname,
        sharePointSiteName: formValue.sharePointSiteName,
        baseFolderPath: formValue.baseFolderPath,
      })
      .subscribe({
        next: result => {
          this.testConnectionResult = result;
          this.testingConnection = false;
        },
        error: () => {
          this.testConnectionResult = {
            success: false,
            message: "Could not reach the server to test the connection.",
          };
          this.testingConnection = false;
        },
      });
  }

  private handleValid(formValue: SharePointPluginConfig): void {
    const valid = !!(
      formValue.configurationTitle &&
      formValue.tenantId &&
      formValue.clientId &&
      formValue.clientSecret &&
      formValue.hostname &&
      formValue.sharePointSiteName &&
      formValue.baseFolderPath
    );
    this.valid$.next(valid);
    this.valid.emit(valid);
  }

  private openSaveSubscription(): void {
    this.saveSubscription = this.save$
      ?.pipe(switchMap(() => combineLatest([this.formValue$, this.valid$]).pipe(take(1))))
      .subscribe(([formValue, valid]) => {
        if (valid) {
          this.configuration.emit(formValue!);
        }
      });
  }
}
