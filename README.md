<div class="eca-actions">
  <button mat-flat-button color="primary" [routerLink]="['/']">
    <mat-icon>arrow_back</mat-icon>Back
  </button>
</div>

<div class="title">
  <h1>Docline Holdings Processor</h1>
</div>

<div class="help-section">

  <h2>Overview</h2>
  <p>
    This Cloud App allows you to update Docline holdings using Alma Analytics data.
    It compares your current Docline holdings against Alma-derived holdings and
    generates properly formatted files for Docline upload.
  </p>

  <h2>Required Inputs</h2>

  <h3>1. Alma Analytics CSV</h3>
  <ul>
    <li>Export from Alma Analytics</li>
    <li>Must include ISSN, Title, MMS ID, and coverage data</li>
    <li>You may upload multiple Analytics files</li>
  </ul>

  <h3>2. Current Docline Holdings CSV</h3>
  <ul>
    <li>Export from Docline</li>
    <li>Must include ISSNs and coverage ranges</li>
  </ul>

  <h2>What the App Does</h2>
  <ul>
    <li>Matches Analytics records to Docline by ISSN</li>
    <li>Builds HOLDING and RANGE records</li>
    <li>Normalizes and merges coverage data</li>
    <li>Compares Alma vs Docline holdings</li>
    <li>Outputs categorized files for Docline upload</li>
  </ul>

  <h2>Output Files</h2>
  <ul>
    <li><strong>Add Final</strong> – New holdings to add</li>
    <li><strong>Update Final</strong> – DELETE + ADD replacements</li>
    <li><strong>Full Match</strong> – No changes required</li>
    <li><strong>Different Ranges</strong> – Coverage differences</li>
    <li><strong>Delete Final</strong> – Remove from Docline</li>
    <li><strong>No Dates</strong> – Missing coverage data</li>
    <li><strong>Counts</strong> – Summary statistics</li>
  </ul>

  <h2>Update Behavior</h2>
  <p>
    Updates are structured as DELETE followed by ADD rows for each record.
    These are automatically sorted to ensure proper Docline processing.
  </p>

  <h2>Settings</h2>
  <ul>
    <li>LIBID</li>
    <li>Retention policy</li>
    <li>Limited retention settings</li>
    <li>EPUB ahead of print</li>
    <li>Supplements</li>
    <li>Ignore warnings</li>
  </ul>

  <p>
    These values are applied to all generated records.
  </p>

  <h2>How to Use</h2>
  <ol>
    <li>Upload Analytics CSV file(s)</li>
    <li>Upload current Docline holdings CSV</li>
    <li>Configure settings</li>
    <li>Click <strong>Process Files</strong></li>
    <li>Download the ZIP output</li>
    <li>Upload appropriate files into Docline</li>
  </ol>

  <h2>Notes</h2>
  <ul>
    <li>Best performance with smaller batches (~50–75 records)</li>
    <li>Matching is driven by ISSN</li>
    <li>All processing occurs in your browser</li>
  </ul>

</div>