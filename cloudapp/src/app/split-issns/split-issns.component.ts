import { Component, OnInit } from '@angular/core';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { CloudAppSettingsService } from '@exlibris/exl-cloudapp-angular-lib';
import { AppSettings, DoclineConfig } from '../models/settings.model';

@Component({
  selector: 'app-split-issns',
  templateUrl: './split-issns.component.html',
  styleUrls: ['./split-issns.component.scss']
})
export class SplitIssnsComponent implements OnInit {
  analyticsFiles: File[] = [];
  doclineFiles: File[] = [];

  loading = false;
  statusMessage = '';
  downloadUrl = '';

  analyticsRows: any[] = [];
  doclineRows: any[] = [];
  mergedRows: any[] = [];
  convertedRows: any[] = [];
  finalRows: any[] = [];
  coverageParseErrorRows: any[] = [];

  previewColumns: string[] = [];
  config: AppSettings = this.getDefaultConfig();

  doclineConfig: DoclineConfig = {
    libid: '',
    retention_policy: 'Permanently Retained',
    limited_retention_period: 0,
    limited_retention_type: 'Years',
    has_epub_ahead_of_print: 'No',
    has_supplements: 'No',
    ignore_warnings: 'No'
  };

  constructor(private settingsService: CloudAppSettingsService) {}

  ngOnInit(): void {
    this.loading = true;

    this.settingsService.get().subscribe({
      next: (savedSettings: any) => {
        this.config = {
          ...this.getDefaultConfig(),
          ...savedSettings,
          doclineConfig: {
            ...this.getDefaultConfig().doclineConfig,
            ...(savedSettings && savedSettings.doclineConfig ? savedSettings.doclineConfig : {})
          }
        };

        this.doclineConfig = {
          ...this.doclineConfig,
          ...this.config.doclineConfig
        };
      },
      error: (err: any) => {
        console.error(err);
      },
      complete: () => {
        this.loading = false;
      }
    });
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

  private readonly DOCLINE_COLUMNS: string[] = [
    'action',
    'record_type',
    'serial_title',
    'nlm_unique_id',
    'holdings_format',
    'begin_volume',
    'end_volume',
    'begin_year',
    'end_year',
    'issns',
    'currently_received',
    'retention_policy',
    'limited_retention_period',
    'limited_retention_type',
    'embargo_period',
    'has_epub_ahead_of_print',
    'has_supplements',
    'ignore_warnings',
    'last_modified'
  ];

  onSelectAnalytics(event: any): void {
    if (event && event.addedFiles && event.addedFiles.length > 0) {
      this.analyticsFiles = this.analyticsFiles.concat(event.addedFiles);
    }
  }

  onRemoveAnalytics(file: File): void {
    this.analyticsFiles = this.analyticsFiles.filter(f => f !== file);
  }

  onSelectDocline(event: any): void {
    if (event && event.addedFiles && event.addedFiles.length > 0) {
      this.doclineFiles = [event.addedFiles[0]];
    }
  }

  onRemoveDocline(file: File): void {
    this.doclineFiles = this.doclineFiles.filter(f => f !== file);
  }

  private downloadZipLocally(zipBlob: Blob, fileName: string): void {
    const url = window.URL.createObjectURL(zipBlob);
    const a = document.createElement('a');

    a.href = url;
    a.download = fileName;
    a.style.display = 'none';

    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);

    window.URL.revokeObjectURL(url);
  }

  async handleUpload(): Promise<void> {
    if (!this.analyticsFiles.length || !this.doclineFiles.length) {
      this.statusMessage = 'Please select one or more Analytics CSVs and one current Docline holdings CSV.';
      return;
    }

    this.loading = true;
    this.statusMessage = 'Reading CSV files...';
    this.downloadUrl = '';
    this.analyticsRows = [];
    this.doclineRows = [];
    this.mergedRows = [];
    this.convertedRows = [];
    this.finalRows = [];
    this.previewColumns = [];
    this.coverageParseErrorRows = [];

    try {
      let analyticsRawRows: any[] = [];

      for (const file of this.analyticsFiles) {
        const rows = await this.readCsvFile(file);
        analyticsRawRows = analyticsRawRows.concat(rows);
      }

      const doclineRawRows = await this.readCsvFile(this.doclineFiles[0]);

      this.statusMessage = 'Normalizing Analytics rows...';
      this.analyticsRows = this.normalizeAnalyticsRows(analyticsRawRows);

      this.statusMessage = 'Normalizing current Docline holdings rows...';
      this.doclineRows = this.normalizeCurrentDoclineRows(doclineRawRows);

      this.statusMessage = 'Propagating current Docline HOLDING values...';
      const propagatedCurrentDoclineRows = this.propagateHoldingValues(this.doclineRows);

      const doclineHoldingRows = propagatedCurrentDoclineRows.filter((row: any) =>
        this.safeString(row['record_type']) === 'HOLDING'
      );

      const explodedDoclineHoldingIssns = this.explodeDoclineIssns(doclineHoldingRows);

      this.statusMessage = 'Grouping Analytics rows...';
      const groupedAnalytics = this.groupAnalyticsRows(this.analyticsRows);

      this.statusMessage = 'Exploding Analytics ISSNs...';
      const explodedAnalyticsIssns = this.explodeAnalyticsIssns(groupedAnalytics);

      this.statusMessage = 'Building Docline ISSN lookup...';
      const doclineIssnIndex = this.buildIndex(explodedDoclineHoldingIssns, 'issn_exploded');

      this.statusMessage = 'Merging Analytics rows to current Docline holdings by ISSN...';
      this.mergedRows = await this.innerJoinChunked(
        explodedAnalyticsIssns.filter((row: any) => this.hasValue(row['ISSN'])),
        'ISSN',
        doclineIssnIndex,
        'issn_exploded',
        'ISSN'
      );

      this.mergedRows = this.dropDuplicatesByKeys(
        this.mergedRows,
        ['nlm_unique_id', 'Electronic or Physical']
      );

      const inferredChoice = this.inferChoiceFromAnalytics(this.analyticsRows);

      this.statusMessage = 'Converting merged rows into Docline HOLDING/RANGE rows...';
      this.convertedRows = this.convertToDoclineRows(this.mergedRows, inferredChoice);

      this.statusMessage = 'Propagating HOLDING values into generated RANGE rows...';
      const propagatedConvertedRows = this.propagateHoldingValues(this.convertedRows);

      this.statusMessage = 'Merging overlapping RANGE intervals...';
      this.finalRows = this.mergeIntervalsOptimized(propagatedConvertedRows);

      if (this.finalRows.length > 0) {
        this.previewColumns = Object.keys(this.finalRows[0]);
      }

      const normalizedCurrentDoclineCompareRows = this.normalizeCompareRows(
        this.propagateHoldingValues(this.doclineRows)
      );

      const normalizedCurrentAlmaCompareRows = this.normalizeCompareRows(this.finalRows);

      this.statusMessage = 'Classifying output sets...';

      const classified = this.classifyOutputSets(
        normalizedCurrentAlmaCompareRows,
        normalizedCurrentDoclineCompareRows,
        inferredChoice
      );

      this.statusMessage = 'Building ZIP package...';

      const zipBlob = await this.buildOutputZip(
        inferredChoice,
        {
          ...classified,
          coverageParseErrorRows: this.coverageParseErrorRows
        }
      );

      this.statusMessage = 'Downloading ZIP package...';
      this.downloadZipLocally(zipBlob, `${inferredChoice}_Docline_Output.zip`);

      this.statusMessage = 'ZIP created and download triggered.';
    } catch (error) {
      console.error(error);
      this.statusMessage = 'Error processing files. Check console.';
    } finally {
      this.loading = false;
    }
  }

  private inferChoiceFromAnalytics(rows: any[]): 'Electronic' | 'Print' {
    const hasCoverage = rows.some((row: any) =>
      this.hasValue(row['Coverage Information Combined'])
    );

    return hasCoverage ? 'Electronic' : 'Print';
  }

  private buildCoverageParseErrorRow(row: any, coverage: string, holdingsFormat: string): any {
    return {
      action: '',
      record_type: 'ERROR',
      serial_title: row['Title_x'],
      nlm_unique_id: row['NLM_Unique_ID'],
      holdings_format: holdingsFormat,
      begin_volume: '',
      end_volume: '',
      begin_year: '',
      end_year: '',
      issns: this.hasValue(row['docline_issns_full'])
        ? this.safeString(row['docline_issns_full'])
        : this.safeString(row['ISSN_x']).replace(/;/g, ','),
      currently_received: '',
      retention_policy: this.doclineConfig.retention_policy,
      limited_retention_period: this.doclineConfig.limited_retention_period,
      limited_retention_type: this.doclineConfig.limited_retention_type,
      embargo_period: 0,
      has_epub_ahead_of_print: this.doclineConfig.has_epub_ahead_of_print,
      has_supplements: this.doclineConfig.has_supplements,
      ignore_warnings: this.doclineConfig.ignore_warnings,
      last_modified: '',
      coverage_statement: coverage,
      error_message: 'Could not derive any non-missing print date range from holdings summary'
    };
  }

  private async readCsvFile(file: File): Promise<any[]> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = () => {
        try {
          const csvText = reader.result as string;
          const workbook = XLSX.read(csvText, {
            type: 'string',
            raw: false
          });

          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];

          let rows: any[] = XLSX.utils.sheet_to_json(worksheet, {
            defval: '',
            raw: false
          });

          rows = rows.map((row: any) => this.trimRowKeys(row));
          resolve(rows);
        } catch (error) {
          reject(error);
        }
      };

      reader.onerror = reject;
      reader.readAsText(file);
    });
  }

  private trimRowKeys(row: any): any {
    const updatedRow: any = {};

    Object.keys(row).forEach((key: string) => {
      updatedRow[key.trim()] = row[key];
    });

    return updatedRow;
  }

  private normalizeAnalyticsRows(rows: any[]): any[] {
    return rows.map((row: any) => {
      const out = this.cloneRow(row);

      out['Electronic or Physical'] = this.safeString(out['Electronic or Physical']);

      if (out['Electronic or Physical'] === 'Physical') {
        out['Electronic or Physical'] = 'Print';
      }

      out['Title'] = this.safeString(out['Title'])
        .replace(/\.$/, '')
        .replace(/^(.+?)(\s\:|$).*/, '$1');

      if (out['Electronic or Physical'] === 'Print') {
        out['Embargo Months'] = '';
        out['Embargo Years'] = '';
      } else {
        out['Embargo Months'] = this.hasValue(out['Embargo Months']) ? out['Embargo Months'] : 0;
        out['Embargo Years'] = this.hasValue(out['Embargo Years']) ? out['Embargo Years'] : 0;
      }

      out['ISSN'] = this.safeString(out['ISSN']);
      out['MMS Id'] = this.safeString(out['MMS Id']);
      out['Network Number (OCoLC)'] = this.safeString(out['Network Number (OCoLC)']).replace(/\.0$/, '');

      return out;
    });
  }

  private normalizeCurrentDoclineRows(rows: any[]): any[] {
    return rows.map((row: any) => {
      const out = this.cloneRow(row);

      out['action'] = this.safeString(out['action']);
      out['record_type'] = this.safeString(out['record_type']);
      out['serial_title'] = this.safeString(out['serial_title']);
      out['nlm_unique_id'] = this.safeString(out['nlm_unique_id']);
      out['holdings_format'] = this.safeString(out['holdings_format']);
      out['issns'] = this.safeString(out['issns']);
      out['begin_volume'] = this.safeString(out['begin_volume']);
      out['end_volume'] = this.safeString(out['end_volume']);
      out['begin_year'] = this.safeString(out['begin_year']);
      out['end_year'] = this.safeString(out['end_year']);
      out['currently_received'] = this.safeString(out['currently_received']);
      out['retention_policy'] = this.safeString(out['retention_policy']);
      out['limited_retention_period'] = this.safeString(out['limited_retention_period']);
      out['limited_retention_type'] = this.safeString(out['limited_retention_type']);
      out['embargo_period'] = this.safeString(out['embargo_period']);
      out['has_epub_ahead_of_print'] = this.safeString(out['has_epub_ahead_of_print']);
      out['has_supplements'] = this.safeString(out['has_supplements']);
      out['ignore_warnings'] = this.safeString(out['ignore_warnings']);
      out['last_modified'] = this.safeString(out['last_modified']);
      out['libid'] = this.safeString(out['libid']);

      return out;
    });
  }

  private groupAnalyticsRows(rows: any[]): any[] {
    const groupedMap: { [key: string]: any } = {};

    rows.forEach((row: any) => {
      const key = [
        this.safeString(row['Title']),
        this.safeString(row['MMS Id']),
        this.safeString(row['ISSN']),
        this.safeString(row['Lifecycle']),
        this.safeString(row['Electronic or Physical'])
      ].join('||');

      if (!groupedMap[key]) {
        groupedMap[key] = this.cloneRow(row);

        if (row.hasOwnProperty('Coverage Information Combined')) {
          groupedMap[key]['Coverage Information Combined'] = this.safeString(row['Coverage Information Combined']);
        }

        if (row.hasOwnProperty('Summary Holdings')) {
          groupedMap[key]['Summary Holdings'] = this.safeString(row['Summary Holdings']);
        }

        groupedMap[key]['Embargo Months'] = [];
        groupedMap[key]['Embargo Years'] = [];
      } else {
        if (this.hasValue(row['Coverage Information Combined'])) {
          groupedMap[key]['Coverage Information Combined'] =
            this.joinWithSemicolon(
              groupedMap[key]['Coverage Information Combined'],
              row['Coverage Information Combined']
            );
        }

        if (this.hasValue(row['Summary Holdings'])) {
          groupedMap[key]['Summary Holdings'] =
            this.joinWithSemicolon(
              groupedMap[key]['Summary Holdings'],
              row['Summary Holdings']
            );
        }
      }

      groupedMap[key]['Embargo Months'].push(
        this.hasValue(row['Embargo Months']) ? row['Embargo Months'] : 0
      );

      groupedMap[key]['Embargo Years'].push(
        this.hasValue(row['Embargo Years']) ? row['Embargo Years'] : 0
      );
    });

    return Object.keys(groupedMap).map((key: string) => {
      const row = groupedMap[key];
      row['Embargo Months'] = JSON.stringify(row['Embargo Months']);
      row['Embargo Years'] = JSON.stringify(row['Embargo Years']);
      return row;
    });
  }

  private explodeAnalyticsIssns(rows: any[]): any[] {
    const output: any[] = [];

    rows.forEach((row: any) => {
      const raw = this.safeString(row['ISSN']);
      const values = raw
        .split(/[;,]/)
        .map((v: string) => v.replace(/\s*\((Print|Electronic)\)\s*/gi, '').trim())
        .filter((v: string) => v.length > 0 && v !== 'None');

      values.forEach((value: string) => {
        const newRow = this.cloneRow(row);
        newRow['ISSN'] = value;
        output.push(newRow);
      });
    });

    return output;
  }

  private explodeDoclineIssns(rows: any[]): any[] {
    const output: any[] = [];

    rows.forEach((row: any) => {
      const raw = this.safeString(row['issns']);

      raw.split(',').forEach((token: string) => {
        const trimmed = token.trim();

        if (!trimmed) {
          return;
        }

        let issnType = '';

        if (/\(Electronic\)/i.test(trimmed)) {
          issnType = 'Electronic';
        } else if (/\(Print\)/i.test(trimmed)) {
          issnType = 'Print';
        }

        const strippedIssn = trimmed
          .replace(/\s*\((Print|Electronic)\)\s*/gi, '')
          .trim();

        if (!strippedIssn) {
          return;
        }

        const newRow = this.cloneRow(row);
        newRow['issn_exploded'] = strippedIssn;
        newRow['issn_label'] = issnType;

        output.push(newRow);
      });
    });

    return output;
  }

  private normalizeCompareRows(rows: any[]): any[] {
    return rows.map((row: any) => {
      const out = this.cloneRow(row);

      out['begin_year'] = this.normalizeCompareYear(out['begin_year']);
      out['end_year'] = this.normalizeCompareYear(out['end_year']);
      out['begin_volume'] = this.safeString(out['begin_volume']);
      out['end_volume'] = this.safeString(out['end_volume']);
      out['embargo_period'] = this.safeString(out['embargo_period']);
      out['currently_received'] = this.safeString(out['currently_received']);
      out['record_type'] = this.safeString(out['record_type']);
      out['holdings_format'] = this.safeString(out['holdings_format']);
      out['nlm_unique_id'] = this.safeString(out['nlm_unique_id']).replace(/^NLM_/, '');

      return out;
    });
  }

  private normalizeCompareYear(value: any): string {
    const s = this.safeString(value);

    if (!s || s === '<NA>') {
      return '0';
    }

    return s;
  }

  private buildIndex(rows: any[], keyName: string): Map<string, any[]> {
    const index = new Map<string, any[]>();

    rows.forEach((row: any) => {
      const key = this.normalizeKey(row[keyName]);

      if (!key) {
        return;
      }

      if (!index.has(key)) {
        index.set(key, []);
      }

      index.get(key)!.push(row);
    });

    return index;
  }

  private async innerJoinChunked(
    leftRows: any[],
    leftKey: string,
    rightIndex: Map<string, any[]>,
    rightMatchField: string,
    matchSource: string,
    chunkSize: number = 1000
  ): Promise<any[]> {
    const output: any[] = [];

    for (let i = 0; i < leftRows.length; i += chunkSize) {
      const chunk = leftRows.slice(i, i + chunkSize);

      chunk.forEach((leftRow: any) => {
        const key = this.normalizeKey(leftRow[leftKey]);

        if (!key) {
          return;
        }

        const matches = rightIndex.get(key) || [];

        matches.forEach((rightRow: any) => {
          const merged = this.cloneRow(leftRow);

          Object.keys(rightRow).forEach((k: string) => {
            if (merged.hasOwnProperty(k)) {
              merged[k + '_docline'] = rightRow[k];
            } else {
              merged[k] = rightRow[k];
            }
          });

          merged['NLM_Unique_ID'] = merged['nlm_unique_id'];
          merged['ISSN_x'] = merged['ISSN'];
          merged['Title_x'] = merged['Title'];
          merged['docline_issns_full'] = this.safeString(rightRow['issns']);

          merged['_match_source'] = matchSource;
          merged['_matched_on_left'] = leftKey;
          merged['_matched_on_right'] = rightMatchField;

          output.push(merged);
        });
      });

      await this.pause();
    }

    return output;
  }

  private convertToDoclineRows(
    mergedRows: any[],
    choice: 'Electronic' | 'Print'
  ): any[] {
    const output: any[] = [];

    mergedRows.forEach((row: any) => {
      let holdingsFormat = row['Electronic or Physical'];

      if (holdingsFormat === 'Physical') {
        holdingsFormat = 'Print';
      }

      const covCombined =
        choice === 'Print'
          ? this.safeString(row['Summary Holdings'])
          : this.safeString(row['Coverage Information Combined']);

      const mainRow: any = {
        'Bibliographic Lifecycle': row['Lifecycle'],
        action: '',
        record_type: 'HOLDING',
        libid: this.doclineConfig.libid,
        serial_title: row['Title_x'],
        nlm_unique_id: row['NLM_Unique_ID'],
        holdings_format: holdingsFormat,
        begin_volume: null,
        end_volume: null,
        begin_year: '',
        end_year: '',
        issns: this.hasValue(row['docline_issns_full'])
          ? this.safeString(row['docline_issns_full'])
          : this.safeString(row['ISSN_x']).replace(/;/g, ','),
        currently_received: this.computeCurrentlyReceived(holdingsFormat, covCombined),
        retention_policy: this.doclineConfig.retention_policy,
        limited_retention_period: this.doclineConfig.limited_retention_period,
        limited_retention_type: this.doclineConfig.limited_retention_type,
        embargo_period: 0,
        has_epub_ahead_of_print: this.doclineConfig.has_epub_ahead_of_print,
        has_supplements: this.doclineConfig.has_supplements,
        ignore_warnings: this.doclineConfig.ignore_warnings,
        last_modified: ''
      };

      output.push(mainRow);

      const coverageEntries = covCombined
        .replace(/;{2,}/g, ';')
        .split(';')
        .map((v: string) => v.trim())
        .filter((v: string) => !!v);

      const embargoMonths = this.parseJsonishArray(row['Embargo Months']);
      const embargoYears = this.parseJsonishArray(row['Embargo Years']);

      const rangeRows: any[] = [];

      coverageEntries.forEach((coverage: string, idx: number) => {
        const month = this.safeInt(embargoMonths[idx], 0);
        const year = this.safeInt(embargoYears[idx], 0);
        const embargoPeriod = choice === 'Electronic' ? (month ? month : year * 12) : 0;

        if (holdingsFormat === 'Print' && this.looksLikePhysicalStatement(coverage)) {
          const ranges = this.parsePhysicalRanges(coverage);

          if (!ranges.length) {
            this.coverageParseErrorRows.push(
              this.buildCoverageParseErrorRow(row, coverage, holdingsFormat)
            );
            return;
          }

          ranges.forEach((pair: any) => {
            rangeRows.push({
              'Bibliographic Lifecycle': row['Lifecycle'],
              action: '',
              record_type: 'RANGE',
              libid: this.doclineConfig.libid,
              serial_title: row['Title_x'],
              nlm_unique_id: row['NLM_Unique_ID'],
              holdings_format: holdingsFormat,
              begin_volume: null,
              end_volume: null,
              begin_year: pair.beginYear,
              end_year: pair.endYear,
              issns: this.hasValue(row['docline_issns_full'])
                ? this.safeString(row['docline_issns_full'])
                : this.safeString(row['ISSN_x']).replace(/;/g, ','),
              currently_received: pair.endYear === null ? 'Yes' : 'No',
              retention_policy: this.doclineConfig.retention_policy,
              limited_retention_period: this.doclineConfig.limited_retention_period,
              limited_retention_type: this.doclineConfig.limited_retention_type,
              embargo_period: embargoPeriod,
              has_epub_ahead_of_print: this.doclineConfig.has_epub_ahead_of_print,
              has_supplements: this.doclineConfig.has_supplements,
              ignore_warnings: this.doclineConfig.ignore_warnings,
              last_modified: ''
            });
          });

          return;
        }

        const electronicRange = this.parseElectronicCoverage(coverage);

        if (electronicRange.beginYear) {
          rangeRows.push({
            'Bibliographic Lifecycle': row['Lifecycle'],
            action: '',
            record_type: 'RANGE',
            libid: this.doclineConfig.libid,
            serial_title: row['Title_x'],
            nlm_unique_id: row['NLM_Unique_ID'],
            holdings_format: holdingsFormat,
            begin_volume: electronicRange.beginVolume,
            end_volume: electronicRange.endVolume,
            begin_year: electronicRange.beginYear,
            end_year: electronicRange.endYear,
            issns: this.hasValue(row['docline_issns_full'])
              ? this.safeString(row['docline_issns_full'])
              : this.safeString(row['ISSN_x']).replace(/;/g, ','),
            currently_received: /until/i.test(covCombined) ? 'No' : 'Yes',
            retention_policy: this.doclineConfig.retention_policy,
            limited_retention_period: this.doclineConfig.limited_retention_period,
            limited_retention_type: this.doclineConfig.limited_retention_type,
            embargo_period: embargoPeriod,
            has_epub_ahead_of_print: this.doclineConfig.has_epub_ahead_of_print,
            has_supplements: this.doclineConfig.has_supplements,
            ignore_warnings: this.doclineConfig.ignore_warnings,
            last_modified: ''
          });
        }
      });

      Array.prototype.push.apply(output, this.dedupeObjects(rangeRows));
    });

    return output;
  }

  private propagateHoldingValues(rows: any[]): any[] {
    const columnsToUpdate = [
      'nlm_unique_id',
      'serial_title',
      'holdings_format',
      'issns',
      'currently_received',
      'retention_policy',
      'limited_retention_period',
      'limited_retention_type',
      'embargo_period',
      'has_epub_ahead_of_print',
      'has_supplements',
      'ignore_warnings',
      'last_modified',
      'libid'
    ];

    const output = rows.map((row: any) => this.cloneRow(row));
    const storedValues: any = {};

    columnsToUpdate.forEach((col: string) => {
      storedValues[col] = null;
    });

    output.forEach((row: any) => {
      if (
        this.safeString(row['record_type']) === 'HOLDING' &&
        this.hasValue(row['nlm_unique_id'])
      ) {
        columnsToUpdate.forEach((col: string) => {
          storedValues[col] = row[col];
        });
      }

      if (this.safeString(row['record_type']) === 'RANGE') {
        columnsToUpdate.forEach((col: string) => {
          if (!this.hasValue(row[col])) {
            row[col] = storedValues[col];
          }
        });
      }
    });

    return output;
  }

  private mergeIntervalsOptimized(rows: any[]): any[] {
    const sortedRows = rows
      .map((row: any) => this.cloneRow(row))
      .sort((a: any, b: any) => {
        const aKey = [
          this.safeString(a['nlm_unique_id']),
          this.safeString(a['holdings_format']),
          this.safeString(a['record_type']),
          this.safeNumberForSort(a['begin_year']),
          this.safeNumberForSort(a['end_year']),
          this.safeNumberForSort(a['embargo_period'])
        ].join('||');

        const bKey = [
          this.safeString(b['nlm_unique_id']),
          this.safeString(b['holdings_format']),
          this.safeString(b['record_type']),
          this.safeNumberForSort(b['begin_year']),
          this.safeNumberForSort(b['end_year']),
          this.safeNumberForSort(b['embargo_period'])
        ].join('||');

        return aKey.localeCompare(bKey);
      });

    sortedRows.forEach((row: any) => {
      if (row['record_type'] === 'RANGE' && !this.hasValue(row['end_year'])) {
        row['end_year'] = 10000;
      }
    });

    const outputRows: any[] = [];
    let currentRow: any = null;

    const effectiveEnd = (row: any): number => {
      const embargo = this.safeInt(row['embargo_period'], 0);
      const endYear = this.safeInt(row['end_year'], 0);
      return endYear - (embargo / 12.0);
    };

    sortedRows.forEach((row: any) => {
      if (row['record_type'] === 'HOLDING') {
        outputRows.push(row);
        return;
      }

      if (!currentRow) {
        currentRow = row;
        return;
      }

      if (
        row['nlm_unique_id'] !== currentRow['nlm_unique_id'] ||
        row['holdings_format'] !== currentRow['holdings_format']
      ) {
        outputRows.push(currentRow);
        currentRow = row;
        return;
      }

      const overlapping =
        this.safeInt(row['begin_year'], 0) <= this.safeInt(currentRow['end_year'], 0) + 1;

      const leftEff = effectiveEnd(row);
      const rightEff = effectiveEnd(currentRow);

      if (this.safeInt(row['begin_year'], 0) > this.safeInt(currentRow['end_year'], 0)) {
        outputRows.push(currentRow);
        currentRow = row;
        return;
      }

      if (overlapping && this.safeInt(row['end_year'], 0) === 10000) {
        currentRow['end_year'] = 10000;
        currentRow['embargo_period'] = row['embargo_period'];
        return;
      }

      if (
        overlapping &&
        this.safeInt(row['embargo_period'], 0) === this.safeInt(currentRow['embargo_period'], 0)
      ) {
        currentRow['end_year'] = Math.max(
          this.safeInt(currentRow['end_year'], 0),
          this.safeInt(row['end_year'], 0)
        );
        return;
      }

      if (overlapping) {
        if (leftEff > rightEff) {
          currentRow['end_year'] = row['end_year'];
          currentRow['embargo_period'] = row['embargo_period'];
        } else if (Math.abs(leftEff - rightEff) < 0.0001) {
          if (
            this.safeInt(row['embargo_period'], 0) <
            this.safeInt(currentRow['embargo_period'], 0)
          ) {
            currentRow['end_year'] = row['end_year'];
            currentRow['embargo_period'] = row['embargo_period'];
          }
        }
        return;
      }

      outputRows.push(currentRow);
      currentRow = row;
    });

    if (currentRow) {
      outputRows.push(currentRow);
    }

    outputRows.forEach((row: any) => {
      if (row['record_type'] === 'RANGE' && this.safeInt(row['end_year'], 0) === 10000) {
        row['currently_received'] = 'Yes';
        row['end_year'] = '';
      }
    });

    return outputRows;
  }

  private classifyOutputSets(
    currentAlmaCompressedRows: any[],
    existingDoclineForCompareRows: any[],
    choice: 'Electronic' | 'Print'
  ): {
    addRows: any[];
    fullMatchRows: any[];
    updateRows: any[];
    differentRangesDoclineRows: any[];
    differentRangesAlmaRows: any[];
    deletedRows: any[];
    inDoclineOnlyPreserveRows: any[];
    noDatesRows: any[];
    countsRows: any[];
  } {
    const almaByKey = this.groupByHoldingKey(currentAlmaCompressedRows);
    const doclineByKey = this.groupByHoldingKey(existingDoclineForCompareRows);

    const allKeys = new Set<string>();

    Object.keys(almaByKey).forEach((key: string) => allKeys.add(key));
    Object.keys(doclineByKey).forEach((key: string) => allKeys.add(key));

    const addRows: any[] = [];
    const fullMatchRows: any[] = [];
    const updateRows: any[] = [];
    const differentRangesDoclineRows: any[] = [];
    const differentRangesAlmaRows: any[] = [];
    const deletedRows: any[] = [];
    const inDoclineOnlyPreserveRows: any[] = [];
    const noDatesRows: any[] = [];

    allKeys.forEach((key: string) => {
      const almaRows = almaByKey[key] || [];
      const doclineRows = doclineByKey[key] || [];

      if (almaRows.length > 0 && doclineRows.length === 0) {
        Array.prototype.push.apply(addRows, almaRows);
        return;
      }

      if (almaRows.length === 0 && doclineRows.length > 0) {
        const shouldDelete = this.shouldDeleteDoclineOnlyRows(doclineRows, choice);

        if (shouldDelete) {
          Array.prototype.push.apply(deletedRows, doclineRows);
        } else {
          Array.prototype.push.apply(inDoclineOnlyPreserveRows, doclineRows);
        }

        return;
      }

      const almaSignature = this.buildRangeSignature(almaRows);
      const doclineSignature = this.buildRangeSignature(doclineRows);

      const issnsCompatible =
        almaRows.length === doclineRows.length &&
        almaRows.every((almaRow: any, idx: number) => {
          const doclineRow = doclineRows[idx];
          return this.issnSetsCompatible(almaRow['issns'], doclineRow['issns']);
        });

      if (almaSignature === doclineSignature && issnsCompatible) {
        Array.prototype.push.apply(fullMatchRows, almaRows);
        return;
      }

      Array.prototype.push.apply(differentRangesAlmaRows, almaRows);
      Array.prototype.push.apply(differentRangesDoclineRows, doclineRows);

      const almaMissingDateRanges = almaRows.filter((row: any) =>
        this.safeString(row['record_type']) === 'RANGE' &&
        !this.hasValue(row['begin_year'])
      );

      Array.prototype.push.apply(noDatesRows, almaMissingDateRanges);

      doclineRows.forEach((row: any) => {
        const cloned = this.cloneRow(row);
        cloned['_update_source'] = 'DOCLINE';
        updateRows.push(cloned);
      });

      almaRows.forEach((row: any) => {
        const cloned = this.cloneRow(row);
        cloned['_update_source'] = 'ALMA';
        updateRows.push(cloned);
      });
    });

    this.applyFinalActionAndPrefix(addRows, 'ADD');
    this.applyFinalActionAndPrefix(fullMatchRows, 'ADD');
    this.applyFinalActionAndPrefix(differentRangesAlmaRows, 'ADD');
    this.applyFinalActionAndPrefix(differentRangesDoclineRows, 'DELETE');
    this.applyFinalActionAndPrefix(deletedRows, 'DELETE');
    this.applyFinalActionAndPrefix(inDoclineOnlyPreserveRows, '');

    updateRows.forEach((row: any) => {
      const nlm = this.safeString(row['nlm_unique_id']).replace(/^NLM_/, '');

      if (nlm) {
        row['nlm_unique_id'] = 'NLM_' + nlm;
      }

      if (row['_update_source'] === 'DOCLINE') {
        row['action'] = 'DELETE';
      } else {
        row['action'] = 'ADD';

        if (this.safeString(row['end_year']) === '0' || row['end_year'] === 0) {
          row['end_year'] = '';
        }
      }

      delete row['_update_source'];
    });

    updateRows.sort((a: any, b: any) => {
      const nlmCompare = this.safeString(a['nlm_unique_id']).localeCompare(
        this.safeString(b['nlm_unique_id'])
      );

      if (nlmCompare !== 0) {
        return nlmCompare;
      }

      const formatCompare = this.safeString(a['holdings_format']).localeCompare(
        this.safeString(b['holdings_format'])
      );

      if (formatCompare !== 0) {
        return formatCompare;
      }

      const actionRank = (action: string): number => {
        if (action === 'DELETE') {
          return 0;
        }

        if (action === 'ADD') {
          return 1;
        }

        return 2;
      };

      const actionCompare =
        actionRank(this.safeString(a['action'])) -
        actionRank(this.safeString(b['action']));

      if (actionCompare !== 0) {
        return actionCompare;
      }

      const recordTypeRank = (recordType: string): number => {
        if (recordType === 'HOLDING') {
          return 0;
        }

        if (recordType === 'RANGE') {
          return 1;
        }

        return 2;
      };

      const recordTypeCompare =
        recordTypeRank(this.safeString(a['record_type'])) -
        recordTypeRank(this.safeString(b['record_type']));

      if (recordTypeCompare !== 0) {
        return recordTypeCompare;
      }

      const beginYearCompare = this.safeString(a['begin_year']).localeCompare(
        this.safeString(b['begin_year'])
      );

      if (beginYearCompare !== 0) {
        return beginYearCompare;
      }

      return this.safeString(a['end_year']).localeCompare(
        this.safeString(b['end_year'])
      );
    });

    this.sortFinalRows(addRows);
    this.sortFinalRows(fullMatchRows);
    this.sortFinalRows(differentRangesAlmaRows);
    this.sortFinalRows(differentRangesDoclineRows);
    this.sortFinalRows(deletedRows);
    this.sortFinalRows(inDoclineOnlyPreserveRows);
    this.sortFinalRows(noDatesRows);

    const normalized = {
      addRows: this.normalizeForFinalOutput(addRows),
      fullMatchRows: this.normalizeForFinalOutput(fullMatchRows),
      updateRows: this.normalizeForFinalOutput(updateRows),
      differentRangesDoclineRows: this.normalizeForFinalOutput(differentRangesDoclineRows),
      differentRangesAlmaRows: this.normalizeForFinalOutput(differentRangesAlmaRows),
      deletedRows: this.normalizeForFinalOutput(deletedRows),
      inDoclineOnlyPreserveRows: this.normalizeForFinalOutput(inDoclineOnlyPreserveRows),
      noDatesRows: this.normalizeForFinalOutput(noDatesRows)
    };

    return {
      ...normalized,
      countsRows: this.buildCountsRowsFromSets(normalized)
    };
  }

  private groupByHoldingKey(rows: any[]): { [key: string]: any[] } {
    const grouped: { [key: string]: any[] } = {};

    rows.forEach((row: any) => {
      const nlm = this.safeString(row['nlm_unique_id']).replace(/^NLM_/, '');
      const format = this.safeString(row['holdings_format']);

      const key = [nlm, format].join('||');

      if (!grouped[key]) {
        grouped[key] = [];
      }

      grouped[key].push(this.cloneRow(row));
    });

    return grouped;
  }

  private issnSetsCompatible(almaIssns: any, doclineIssns: any): boolean {
    const normalize = (value: any): string[] => {
      return this.safeString(value)
        .split(',')
        .map((v: string) => v.replace(/\s*\((Print|Electronic)\)\s*/gi, '').trim())
        .filter((v: string) => !!v)
        .sort();
    };

    const alma = normalize(almaIssns);
    const docline = normalize(doclineIssns);

    if (!alma.length && !docline.length) {
      return true;
    }

    for (let i = 0; i < alma.length; i++) {
      if (docline.indexOf(alma[i]) === -1) {
        return false;
      }
    }

    return true;
  }

  private normalizeIssnsForCompare(value: any): string {
    const raw = this.safeString(value);

    if (!raw) {
      return '';
    }

    const parts = raw
      .split(',')
      .map((v: string) => v.replace(/\s*\((Print|Electronic)\)\s*/gi, '').trim())
      .filter((v: string) => !!v)
      .sort();

    return parts.join(',');
  }

  private buildRangeSignature(rows: any[]): string {
    const normalized = rows
      .map((row: any) => ({
        record_type: this.safeString(row['record_type']),
        begin_volume: this.safeString(row['begin_volume']),
        end_volume: this.safeString(row['end_volume']),
        begin_year: this.normalizeCompareYear(row['begin_year']),
        end_year: this.normalizeCompareYear(row['end_year']),
        embargo_period: this.safeString(row['embargo_period']),
        currently_received: this.safeString(row['currently_received']),
        holdings_format: this.safeString(row['holdings_format']),
        issns: this.normalizeIssnsForCompare(row['issns'])
      }))
      .sort((a: any, b: any) => {
        const aKey = [
          a.record_type,
          a.begin_volume,
          a.end_volume,
          a.begin_year,
          a.end_year,
          a.embargo_period,
          a.currently_received,
          a.holdings_format,
          a.issns
        ].join('||');

        const bKey = [
          b.record_type,
          b.begin_volume,
          b.end_volume,
          b.begin_year,
          b.end_year,
          b.embargo_period,
          b.currently_received,
          b.holdings_format,
          b.issns
        ].join('||');

        return aKey.localeCompare(bKey);
      });

    return JSON.stringify(normalized);
  }

  private shouldDeleteDoclineOnlyRows(
    rows: any[],
    choice: 'Electronic' | 'Print'
  ): boolean {
    const formats = new Set(rows.map((r: any) => this.safeString(r['holdings_format'])));

    if (choice === 'Electronic') {
      return formats.has('Electronic');
    }

    return false;
  }

  private applyFinalActionAndPrefix(rows: any[], action: string): void {
    rows.forEach((row: any) => {
      const nlm = this.safeString(row['nlm_unique_id']).replace(/^NLM_/, '');

      if (nlm) {
        row['nlm_unique_id'] = 'NLM_' + nlm;
      }

      row['action'] = action;

      if (action === 'ADD') {
        if (this.safeString(row['end_year']) === '0' || row['end_year'] === 0) {
          row['end_year'] = '';
        }
      }
    });
  }

  private sortFinalRows(rows: any[]): void {
    rows.sort((a: any, b: any) => {
      const serialTitleCompare = this.safeString(a['serial_title']).localeCompare(
        this.safeString(b['serial_title'])
      );

      if (serialTitleCompare !== 0) {
        return serialTitleCompare;
      }

      const nlmCompare = this.safeString(a['nlm_unique_id']).localeCompare(
        this.safeString(b['nlm_unique_id'])
      );

      if (nlmCompare !== 0) {
        return nlmCompare;
      }

      const actionCompare = this.safeString(b['action']).localeCompare(
        this.safeString(a['action'])
      );

      if (actionCompare !== 0) {
        return actionCompare;
      }

      const recordTypeCompare = this.safeString(a['record_type']).localeCompare(
        this.safeString(b['record_type'])
      );

      if (recordTypeCompare !== 0) {
        return recordTypeCompare;
      }

      const embargoCompare = this.safeString(a['embargo_period']).localeCompare(
        this.safeString(b['embargo_period'])
      );

      if (embargoCompare !== 0) {
        return embargoCompare;
      }

      const beginYearCompare = this.safeString(a['begin_year']).localeCompare(
        this.safeString(b['begin_year'])
      );

      if (beginYearCompare !== 0) {
        return beginYearCompare;
      }

      return this.safeString(a['end_year']).localeCompare(
        this.safeString(b['end_year'])
      );
    });
  }

  private normalizeForFinalOutput(rows: any[]): any[] {
    const dropCols = new Set([
      'libid',
      'index',
      'level_0',
      'Lifecycle',
      'Bibliographic Lifecycle'
    ]);

    return rows.map((row: any) => {
      const cleaned: any = {};

      Object.keys(row).forEach((key: string) => {
        const trimmed = String(key).trim();

        if (!dropCols.has(trimmed)) {
          let value = row[key];

          if (trimmed === 'end_year' && (value === '0' || value === 0)) {
            value = '';
          }

          cleaned[trimmed] = value;
        }
      });

      return cleaned;
    });
  }

  private buildCountsRowsFromSets(outputSets: {
    addRows: any[];
    fullMatchRows: any[];
    updateRows: any[];
    differentRangesDoclineRows: any[];
    differentRangesAlmaRows: any[];
    deletedRows: any[];
    inDoclineOnlyPreserveRows: any[];
    noDatesRows: any[];
  }): any[] {
    const makeRow = (setName: string, rows: any[]) => ({
      Set: setName,
      'Number of Rows': rows.length,
      'Number of NLM Unique IDs': this.uniqueCountByKey(rows, 'nlm_unique_id')
    });

    return [
      makeRow('Alma Adds', outputSets.addRows),
      makeRow('Full Match', outputSets.fullMatchRows),
      makeRow('Alma Updates', outputSets.updateRows),
      makeRow('Docline Different Ranges', outputSets.differentRangesDoclineRows),
      makeRow('Alma Different Ranges', outputSets.differentRangesAlmaRows),
      makeRow('Deleted from Alma', outputSets.deletedRows),
      makeRow('In Docline Only Keep', outputSets.inDoclineOnlyPreserveRows),
      makeRow('No Dates', outputSets.noDatesRows)
    ];
  }

  private projectDoclineColumns(rows: any[]): any[] {
    return rows.map((row: any) => {
      const projected: any = {};

      this.DOCLINE_COLUMNS.forEach((col: string) => {
        let value = row[col];

        if (value === null || value === undefined) {
          value = '';
        }

        projected[col] = value;
      });

      return projected;
    });
  }

  private escapeCsvValue(value: any): string {
    if (value === null || value === undefined) {
      return '';
    }

    const s = String(value);

    if (
      s.indexOf('"') > -1 ||
      s.indexOf(',') > -1 ||
      s.indexOf('\n') > -1 ||
      s.indexOf('\r') > -1
    ) {
      return '"' + s.replace(/"/g, '""') + '"';
    }

    return s;
  }

  private rowsToCsv(rows: any[], columns: string[]): string {
    const header = columns.join(',');

    const lines = rows.map((row: any) => {
      return columns.map((col: string) => this.escapeCsvValue(row[col])).join(',');
    });

    return [header].concat(lines).join('\r\n');
  }

  private uniqueCountByKey(rows: any[], keyName: string): number {
    const set = new Set<string>();

    rows.forEach((row: any) => {
      const value = this.safeString(row[keyName]);

      if (value) {
        set.add(value);
      }
    });

    return set.size;
  }

  private async buildOutputZip(
    choice: 'Electronic' | 'Print',
    outputSets: {
      addRows: any[];
      fullMatchRows: any[];
      updateRows: any[];
      differentRangesDoclineRows: any[];
      differentRangesAlmaRows: any[];
      deletedRows: any[];
      inDoclineOnlyPreserveRows: any[];
      noDatesRows: any[];
      countsRows: any[];
      coverageParseErrorRows: any[];
    }
  ): Promise<Blob> {
    const zip = new JSZip();
    const outputFolder = zip.folder('Output');

    if (!outputFolder) {
      throw new Error('Could not create Output folder in ZIP.');
    }

    outputFolder.file(
      `${choice} Coverage Parse Errors.csv`,
      this.rowsToCsv(
        outputSets.coverageParseErrorRows,
        [
          'action',
          'record_type',
          'serial_title',
          'nlm_unique_id',
          'holdings_format',
          'begin_volume',
          'end_volume',
          'begin_year',
          'end_year',
          'issns',
          'currently_received',
          'retention_policy',
          'limited_retention_period',
          'limited_retention_type',
          'embargo_period',
          'has_epub_ahead_of_print',
          'has_supplements',
          'ignore_warnings',
          'last_modified',
          'coverage_statement',
          'error_message'
        ]
      )
    );

    outputFolder.file(
      `${choice} Add Final.csv`,
      this.rowsToCsv(this.projectDoclineColumns(outputSets.addRows), this.DOCLINE_COLUMNS)
    );

    outputFolder.file(
      `${choice} Full Match Final.csv`,
      this.rowsToCsv(this.projectDoclineColumns(outputSets.fullMatchRows), this.DOCLINE_COLUMNS)
    );

    outputFolder.file(
      `${choice} Update Final.csv`,
      this.rowsToCsv(this.projectDoclineColumns(outputSets.updateRows), this.DOCLINE_COLUMNS)
    );

    outputFolder.file(
      `${choice} Different Ranges Docline Final.csv`,
      this.rowsToCsv(this.projectDoclineColumns(outputSets.differentRangesDoclineRows), this.DOCLINE_COLUMNS)
    );

    outputFolder.file(
      `${choice} Different Ranges Alma Final.csv`,
      this.rowsToCsv(this.projectDoclineColumns(outputSets.differentRangesAlmaRows), this.DOCLINE_COLUMNS)
    );

    outputFolder.file(
      `${choice} Delete Final - Either Withdrawn from Alma or ILL Not Allowed for E-Resources.csv`,
      this.rowsToCsv(this.projectDoclineColumns(outputSets.deletedRows), this.DOCLINE_COLUMNS)
    );

    outputFolder.file(
      `${choice} In Docline Only Preserve Final.csv`,
      this.rowsToCsv(this.projectDoclineColumns(outputSets.inDoclineOnlyPreserveRows), this.DOCLINE_COLUMNS)
    );

    outputFolder.file(
      'No Dates in Update Table.csv',
      this.rowsToCsv(this.projectDoclineColumns(outputSets.noDatesRows), this.DOCLINE_COLUMNS)
    );

    outputFolder.file(
      `${choice} Counts.csv`,
      this.rowsToCsv(
        outputSets.countsRows,
        ['Set', 'Number of Rows', 'Number of NLM Unique IDs']
      )
    );

    return await zip.generateAsync({
      type: 'blob'
    });
  }

  private async uploadZipToServer(
    zipBlob: Blob,
    zipFileName: string
  ): Promise<string> {
    const formData = new FormData();
    formData.append('file', zipBlob, zipFileName);

    const response = await fetch('https://your-server.example.edu/upload-docline-zip', {
      method: 'POST',
      body: formData
    });

    if (!response.ok) {
      throw new Error('Failed to upload ZIP to server.');
    }

    const data = await response.json();

    if (!data || !data.download_url) {
      throw new Error('Server did not return a download URL.');
    }

    return data.download_url;
  }

  private parseElectronicCoverage(coverage: string): any {
    let beginYear: any = null;
    let endYear: any = null;
    let beginVolume: any = null;
    let endVolume: any = null;

    if (/from/i.test(coverage)) {
      if (/until/i.test(coverage)) {
        const beginning = coverage.split(/from/i)[1].split(/until/i)[0].trim();
        const ending = coverage.split(/until/i)[1].trim();

        const beginYearMatch = beginning.match(/(\d{4})/);
        const endYearMatch = ending.match(/(\d{4})/);

        beginYear = beginYearMatch ? beginYearMatch[1] : null;
        endYear = endYearMatch ? endYearMatch[1] : null;

        const beginVolumeMatch = beginning.match(/volume\:\s*(\d+)/i);
        if (beginVolumeMatch) {
          beginVolume = beginVolumeMatch[1];
        }

        const endVolumeMatch = ending.match(/volume\:\s*(\d+)/i);
        if (endVolumeMatch) {
          endVolume = endVolumeMatch[1];
        }
      } else {
        const beginningOnly = coverage.split(/from/i)[1].trim();
        const beginYearMatch = beginningOnly.match(/(\d{4})/);
        beginYear = beginYearMatch ? beginYearMatch[1] : null;

        const beginVolumeMatch = beginningOnly.match(/volume\:\s*(\d+)/i);
        if (beginVolumeMatch) {
          beginVolume = beginVolumeMatch[1];
        }
      }
    } else {
      const year = this.extractYear(coverage);

      if (year) {
        beginYear = year;
        endYear = year;
      }
    }

    return {
      beginYear,
      endYear,
      beginVolume,
      endVolume
    };
  }

  private parsePhysicalRanges(statement: string): any[] {
    if (!statement) {
      return [];
    }

    const segments = statement
      .split(/;\s*/)
      .map((s: string) => s.trim())
      .filter((s: string) => !!s);

    const positiveSegments: string[] = [];
    let inMissingBlock = false;

    segments.forEach((rawSegment: string) => {
      let seg = rawSegment.trim();

      if (/^Missing:\s*/i.test(seg)) {
        inMissingBlock = true;
        seg = seg.replace(/^Missing:\s*/i, '').trim();
      }

      const hasYear = !!this.extractYear(seg);

      if (inMissingBlock) {
        if (!hasYear) {
          return;
        }

        inMissingBlock = false;
      }

      if (seg) {
        positiveSegments.push(seg);
      }
    });

    if (!positiveSegments.length) {
      return [];
    }

    const ranges: any[] = [];

    positiveSegments.forEach((seg: string) => {
      if (seg.indexOf('-') > -1) {
        const pieces = seg.split('-', 2);
        const left = pieces[0];
        const right = pieces[1] || '';

        const beginYear = this.extractYear(left);
        let endYear = this.extractYear(right);

        const openEnded = right.trim() === '' || /-\s*$/.test(seg);
        if (openEnded) {
          endYear = null;
        }

        if (beginYear) {
          ranges.push({
            beginYear,
            endYear
          });
        }
      } else {
        const y = this.extractYear(seg);

        if (y) {
          ranges.push({
            beginYear: y,
            endYear: y
          });
        }
      }
    });

    return this.dedupeByKey(ranges, ['beginYear', 'endYear']);
  }

  private computeCurrentlyReceived(
    holdingsFormat: string,
    coverageCombined: string
  ): string {
    if (holdingsFormat === 'Print' && this.looksLikePhysicalStatement(coverageCombined)) {
      return /-\s*(;|$)/.test(coverageCombined) ? 'Yes' : 'No';
    }

    return /\buntil\b/i.test(coverageCombined) ? 'No' : 'Yes';
  }

  private looksLikePhysicalStatement(s: string): boolean {
    return !!s && (
      s.indexOf('v.') > -1 ||
      /Missing:/i.test(s) ||
      /\(\d{4}/.test(s)
    );
  }

  private extractYear(s: string): string | null {
    if (!s) {
      return null;
    }

    const parenMatch = s.match(/\((\d{4})/);

    if (parenMatch) {
      return parenMatch[1];
    }

    const genericMatch = s.match(/\b(1[6-9]\d{2}|20\d{2})\b/);
    return genericMatch ? genericMatch[1] : null;
  }

  private parseJsonishArray(value: any): any[] {
    const s = this.safeString(value);

    if (!s) {
      return [];
    }

    try {
      const parsed = JSON.parse(s);
      return Array.isArray(parsed) ? parsed : [parsed];
    } catch {
      return s.split(',').map((v: string) => v.trim());
    }
  }

  private dedupeObjects(rows: any[]): any[] {
    const seen = new Set<string>();
    const output: any[] = [];

    rows.forEach((row: any) => {
      const key = JSON.stringify(row);

      if (!seen.has(key)) {
        seen.add(key);
        output.push(row);
      }
    });

    return output;
  }

  private dedupeByKey(rows: any[], keys: string[]): any[] {
    const seen = new Set<string>();
    const output: any[] = [];

    rows.forEach((row: any) => {
      const key = keys.map((k: string) => this.safeString(row[k])).join('||');

      if (!seen.has(key)) {
        seen.add(key);
        output.push(row);
      }
    });

    return output;
  }

  private dropDuplicatesByKeys(rows: any[], keys: string[]): any[] {
    return this.dedupeByKey(rows, keys);
  }

  private pause(): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, 0));
  }

  private cloneRow(row: any): any {
    const out: any = {};

    Object.keys(row).forEach((k: string) => {
      out[k] = row[k];
    });

    return out;
  }

  private safeString(value: any): string {
    if (value === null || value === undefined) {
      return '';
    }

    return String(value).trim();
  }

  private hasValue(value: any): boolean {
    const s = this.safeString(value);
    return s !== '' && s !== 'None' && s !== 'nan' && s !== '<NA>';
  }

  private normalizeKey(value: any): string {
    return this.safeString(value).replace(/\s+/g, ' ').trim();
  }

  private joinWithSemicolon(existingValue: any, nextValue: any): string {
    const left = this.safeString(existingValue);
    const right = this.safeString(nextValue);

    if (!left) {
      return right;
    }

    if (!right) {
      return left;
    }

    return left + ';' + right;
  }

  private safeInt(value: any, defaultValue: number = 0): number {
    if (value === null || value === undefined) {
      return defaultValue;
    }

    const s = String(value)
      .replace(/\[/g, '')
      .replace(/\]/g, '')
      .trim();

    if (!s || s === 'None' || s === 'nan' || s === '<NA>') {
      return defaultValue;
    }

    const parsed = parseFloat(s);
    return isNaN(parsed) ? defaultValue : parsed;
  }

  private safeNumberForSort(value: any): string {
    const n = this.safeInt(value, 0);
    return String(n).padStart(10, '0');
  }
}