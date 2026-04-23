# Docline Holdings Processor

## Overview
This Cloud App allows you to update Docline holdings using Alma Analytics data.  
It compares your current Docline holdings against Alma-derived holdings and generates properly formatted files for Docline upload.

---

## Required Inputs

### 1. Alma Analytics CSV
- Export from Alma Analytics
- Must include ISSN, Title, MMS ID, and coverage data
- You may upload multiple Analytics files

### 2. Current Docline Holdings CSV
- Export from Docline
- Must include ISSNs and coverage ranges

---

## What the App Does
- Matches Analytics records to Docline by ISSN
- Builds HOLDING and RANGE records
- Normalizes and merges coverage data
- Compares Alma vs Docline holdings
- Outputs categorized files for Docline upload

---

## Output Files
- **Add Final** – New holdings to add  
- **Update Final** – DELETE + ADD replacements  
- **Full Match** – No changes required  
- **Different Ranges** – Coverage differences  
- **Delete Final** – Remove from Docline  
- **No Dates** – Missing coverage data  
- **Counts** – Summary statistics  

---

## Update Behavior
Updates are structured as **DELETE followed by ADD** rows for each record.  
These are automatically sorted to ensure proper Docline processing.

---

## Settings
- LIBID
- Retention policy
- Limited retention settings
- EPUB ahead of print
- Supplements
- Ignore warnings

These values are applied to all generated records.

---

## How to Use
1. Upload Analytics CSV file(s)
2. Upload current Docline holdings CSV
3. Configure settings
4. Click **Process Files**
5. Download the ZIP output
6. Upload appropriate files into Docline

---

## Notes
- Best performance with smaller batches (~50–75 records)
- Matching is driven by ISSN
- All processing occurs in your browser
</div>