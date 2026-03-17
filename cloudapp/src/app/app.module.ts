import { NgModule, APP_INITIALIZER } from '@angular/core';
import { HttpClientModule } from '@angular/common/http';
import { BrowserModule } from '@angular/platform-browser';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { MAT_FORM_FIELD_DEFAULT_OPTIONS } from '@angular/material/form-field';
import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import {
  MaterialModule,
  CloudAppTranslateModule,
  AlertModule,
  InitService
} from '@exlibris/exl-cloudapp-angular-lib';
import { NgxDropzoneModule } from 'ngx-dropzone';
import { MatIconModule } from '@angular/material/icon';

import { AppComponent } from './app.component';
import { AppRoutingModule } from './app-routing.module';
import { SplitIssnsComponent } from './split-issns/split-issns.component';
//import { BrowserModule } from '@angular/platform-browser';
import { CommonModule } from '@angular/common';
import { SettingsComponent } from './settings/settings.component';


@NgModule({
  declarations: [
    AppComponent,
    SplitIssnsComponent,
    SettingsComponent
  ],
  imports: [
    MaterialModule,
    BrowserModule,
    CommonModule,
    BrowserAnimationsModule,
    AppRoutingModule,
    HttpClientModule,
    AlertModule,
    FormsModule,
    ReactiveFormsModule,
    NgxDropzoneModule,
    MatIconModule,
    CloudAppTranslateModule.forRoot()
  ],
//   providers: [
//     {
//       provide: APP_INITIALIZER,
//       deps: [InitService],
//       multi: true
//     },
//     { provide: MAT_FORM_FIELD_DEFAULT_OPTIONS, useValue: { appearance: 'standard' } }
//   ],
  bootstrap: [AppComponent]
})
export class AppModule { }