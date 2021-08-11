import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { DecomposerComponent } from './decomposer/decomposer.component';
import { MergerComponent } from './merger/merger.component';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';

import {MatButtonModule} from '@angular/material/button';
import { ColumnListComponent } from './column-list/column-list.component';
import {ScrollingModule} from '@angular/cdk/scrolling';
import { ProgressBarComponent } from './progress-bar/progress-bar.component';
import {MatListModule} from '@angular/material/list';
import {MatProgressBarModule} from '@angular/material/progress-bar';
import {MatCardModule} from '@angular/material/card';
import {MatInputModule} from '@angular/material/input';
import {MatCheckboxModule} from '@angular/material/checkbox';
import { FormsModule } from '@angular/forms';
import {MatRadioModule} from '@angular/material/radio';
@NgModule({
  declarations: [
    AppComponent,
    DecomposerComponent,
    MergerComponent,
    ColumnListComponent,
    ProgressBarComponent
  ],
  imports: [
    BrowserModule,
    AppRoutingModule,
    BrowserAnimationsModule,
    MatButtonModule,
	ScrollingModule,
	MatListModule,
	MatProgressBarModule,
	MatCardModule,
	MatInputModule,
	MatCheckboxModule,
	FormsModule,
    MatButtonModule,
	MatRadioModule
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
