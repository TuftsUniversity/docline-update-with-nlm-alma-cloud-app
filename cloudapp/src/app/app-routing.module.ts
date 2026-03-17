import { NgModule } from '@angular/core';
import { Routes, RouterModule } from '@angular/router';
import { SplitIssnsComponent } from './split-issns/split-issns.component';
import { SettingsComponent } from './settings/settings.component';
const routes: Routes = [

  { path: '', component: SplitIssnsComponent },
  { path: 'settings', component: SettingsComponent }
];

@NgModule({
  imports: [RouterModule.forRoot(routes, { useHash: true })],
  exports: [RouterModule]
})
export class AppRoutingModule { }