import { Component, OnInit } from '@angular/core';
import { NgForm } from '@angular/forms';
import {
  AlertService,
  CloudAppSettingsService,
  RestErrorResponse
} from '@exlibris/exl-cloudapp-angular-lib';
import { Router } from '@angular/router';
import { AppSettings } from '../models/settings.model';

@Component({
  selector: 'app-settings',
  templateUrl: './settings.component.html',
  styleUrls: ['./settings.component.scss']
})
export class SettingsComponent implements OnInit {
  config: AppSettings = this.getDefaultConfig();
  loading = false;

  constructor(
    private settingsService: CloudAppSettingsService,
    private alert: AlertService,
    private router: Router
  ) {}

  ngOnInit(): void {
    this.loading = true;

    this.settingsService.get().subscribe({
      next: (savedConfig: any) => {
        this.config = {
          ...this.getDefaultConfig(),
          ...savedConfig,
          doclineConfig: {
            ...this.getDefaultConfig().doclineConfig,
            ...(savedConfig && savedConfig.doclineConfig ? savedConfig.doclineConfig : {})
          }
        };
      },
      error: (err: any) => {
        console.error(err);
        this.alert.error('Unable to load settings.');
      },
      complete: () => {
        this.loading = false;
      }
    });
  }

  onSubmit(form: NgForm): void {
    if (!this.config.doclineConfig.libid) {
      this.alert.error('LIBID is required.');
      return;
    }

    this.settingsService.set(this.config).subscribe({
      next: () => {
        this.alert.success('Updated Successfully', { keepAfterRouteChange: true });
        this.router.navigate(['']);
      },
      error: (err: RestErrorResponse) => {
        console.error(err.message);
        this.alert.error(err.message);
      }
    });
  }

  onRestore(): void {
    this.config = this.getDefaultConfig();
  }

  private getDefaultConfig(): AppSettings {
    return {
      doclineConfig: {
        libid: '',
        retention_policy: 'Permanently Retained',
        limited_retention_period: 0,
        limited_retention_type: 'Years',
        has_epub_ahead_of_print: 'No',
        has_supplements: 'No',
        ignore_warnings: 'No'
      }
    };
  }
}