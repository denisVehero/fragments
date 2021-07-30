import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { DecomposerComponent } from './decomposer/decomposer.component';
import { MergerComponent } from './merger/merger.component';

@NgModule({
  declarations: [
    AppComponent,
    DecomposerComponent,
    MergerComponent
  ],
  imports: [
    BrowserModule,
    AppRoutingModule
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
