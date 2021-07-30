import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { DecomposerComponent } from './decomposer/decomposer.component';
import { MergerComponent } from './merger/merger.component';
const routes: Routes = [
  {path: "decomposer", component: DecomposerComponent},
  {path: "merger", component: MergerComponent},
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule { }
