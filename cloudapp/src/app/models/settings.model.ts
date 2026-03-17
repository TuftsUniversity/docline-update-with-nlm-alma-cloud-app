export interface DoclineConfig {
  libid: string;
  retention_policy: string;
  limited_retention_period: number;
  limited_retention_type: string;
  has_epub_ahead_of_print: string;
  has_supplements: string;
  ignore_warnings: string;
}

export interface AppSettings {
  doclineConfig: DoclineConfig;
}