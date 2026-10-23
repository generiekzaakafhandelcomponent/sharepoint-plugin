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

import {ComponentFixture, TestBed} from "@angular/core/testing";
import {NO_ERRORS_SCHEMA, Pipe, PipeTransform} from "@angular/core";
import {of, Subject} from "rxjs";
import {DeleteWorkDocumentConfigurationComponent} from "./delete-work-document-configuration.component";

@Pipe({name: "pluginTranslate", standalone: false})
class FakePluginTranslatePipe implements PipeTransform {
  transform(key: string) {
    return of(key);
  }
}

describe("DeleteWorkDocumentConfigurationComponent", () => {
  let fixture: ComponentFixture<DeleteWorkDocumentConfigurationComponent>;
  let component: DeleteWorkDocumentConfigurationComponent;
  const save$ = new Subject<void>();

  beforeEach(() => {
    TestBed.configureTestingModule({
      declarations: [DeleteWorkDocumentConfigurationComponent, FakePluginTranslatePipe],
      schemas: [NO_ERRORS_SCHEMA],
    });
    fixture = TestBed.createComponent(DeleteWorkDocumentConfigurationComponent);
    component = fixture.componentInstance;
    component.save$ = save$;
    component.disabled$ = of(false);
    component.pluginId = "sharepoint-plugin";
    fixture.detectChanges();
  });

  it("is valid only with a file name", () => {
    const valid: boolean[] = [];
    component.valid.subscribe(value => valid.push(value));

    component.formValueChange({fileName: ""});
    component.formValueChange({fileName: "pv:bestandsnaam"});

    expect(valid).toEqual([false, true]);
  });

  it("saves the file name as the action's configuration", () => {
    const configurations: object[] = [];
    component.configuration.subscribe(value => configurations.push(value));

    component.formValueChange({fileName: "pv:bestandsnaam"});
    save$.next();

    expect(configurations).toEqual([{fileName: "pv:bestandsnaam"}]);
  });
});
