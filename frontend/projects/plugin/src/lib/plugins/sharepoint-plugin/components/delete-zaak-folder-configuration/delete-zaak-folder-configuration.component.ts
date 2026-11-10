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
import {FunctionConfigurationComponent, FunctionConfigurationData} from "@valtimo/plugin";
import {Observable, Subscription} from "rxjs";

// The delete action has no properties: always valid, saves an empty configuration.
@Component({
  standalone: false,
  selector: "valtimo-delete-zaak-folder-configuration",
  templateUrl: "./delete-zaak-folder-configuration.component.html",
})
export class DeleteZaakFolderConfigurationComponent
  implements FunctionConfigurationComponent, OnInit, OnDestroy
{
  @Input() save$!: Observable<void>;
  @Input() disabled$!: Observable<boolean>;
  @Input() pluginId!: string;
  @Input() prefillConfiguration$!: Observable<FunctionConfigurationData>;
  @Output() valid: EventEmitter<boolean> = new EventEmitter<boolean>();
  @Output() configuration: EventEmitter<FunctionConfigurationData> =
    new EventEmitter<FunctionConfigurationData>();

  private saveSubscription!: Subscription;

  ngOnInit(): void {
    this.valid.emit(true);
    this.saveSubscription = this.save$?.subscribe(() => this.configuration.emit({}));
  }

  ngOnDestroy(): void {
    this.saveSubscription?.unsubscribe();
  }
}
