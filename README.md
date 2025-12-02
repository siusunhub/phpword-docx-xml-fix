# PHPWord Template XML Fixer

Small helper tool to repair **PHPWord** DOCX templates where placeholders like `${NAME}` get broken in the underlying DOCX XML and then **cannot be replaced** by PHPWord.

Word sometimes splits a placeholder into multiple pieces, e.g.:

```text
${NA    ME}
```

Visually in Word this still looks like `${NAME}`, but internally it’s no longer a single continuous string, so:

```php
$templateProcessor->setValue('NAME', 'John');
```

does nothing.  
This script scans the DOCX XML, finds those broken placeholders, and fixes them.

---

## Why this is needed

Inside a DOCX file, your document is stored as `word/document.xml`.  
A simple placeholder like:

```text
${NAME}
```

can end up split across runs, tags, or even contain a tab:

```xml
<w:p>
  <w:r>
    <w:t>${NA</w:t>
  </w:r>
  <w:r>
    <w:tab/>
  </w:r>
  <w:r>
    <w:t>ME}</w:t>
  </w:r>
</w:p>
```

Or with a literal tab character inside the text node:

```xml
<w:p>
  <w:r>
    <w:t xml:space="preserve">${NA	M</w:t>
  </w:r>
  <w:r>
    <w:t>E}</w:t>
  </w:r>
</w:p>
```

To the user this still shows as `${NAME}`, but for PHPWord it’s no longer the exact `${NAME}` string, so `TemplateProcessor::setValue('NAME', ...)` fails.

**This script:**

- Extracts `document.xml` from the DOCX.
- Searches for broken patterns around `${...}` that span tags, runs, or contain tabs/whitespace.
- Strips the XML tags and unwanted characters *inside* the placeholder.
- Writes everything back so that the XML contains a clean `${NAME}` again.

---

## Features

- Fixes PHPWord-style placeholders of the form `${XXX}` that are split by:
  - `<w:t>`, `<w:tab/>`, or other Word XML tags
  - tabs and stray whitespace inserted by Word
- Keeps the rest of the DOCX untouched.
- Produces a new, fixed DOCX file ready for use with PHPWord’s `TemplateProcessor`.

---

## Requirements

- PHP **7+**
- `ZipArchive` extension enabled
- A web server that can run PHP (e.g. Apache, Nginx + PHP-FPM)

---

## Installation

1. Copy `xmldocumentfix.php` into a directory served by your web server.
2. Make sure the file is accessible in a browser, for example:

   ```text
   http://your-server/xmldocumentfix.php
   ```

3. Ensure PHP has permission to write temporary files to `/tmp` (or adapt the script if you want a different temp folder).

---

## Usage

1. Open the script in your browser:

   ```text
   http://your-server/xmldocumentfix.php
   ```

2. You’ll see a simple form:

   - **Select docx file** – choose your DOCX template with `${...}` placeholders.
   - **Show Debug Only** – process the file but **don’t** send the fixed DOCX to download, just show the log.
   - **Show All Template Tag** – log all detected template tags, not only the ones that were changed.
   - Click **Scan and Fix**.

3. If not in debug-only mode, the script will:

   - Extract `word/document.xml` from the uploaded DOCX.
   - Scan for broken `${XXX}` placeholders (e.g. `${NA    ME}`).
   - Fix them and build a new DOCX.
   - Trigger a download of the fixed DOCX file.

4. Use the downloaded DOCX as your PHPWord template:

   ```php
   use PhpOffice\PhpWord\TemplateProcessor;

   $templateProcessor = new TemplateProcessor('fixed-template.docx');
   $templateProcessor->setValue('NAME', 'John');
   $templateProcessor->saveAs('output.docx');
   ```

---

## How it works (internally)

- Unzips the uploaded DOCX into a temporary folder.
- Moves `word/document.xml` to a working file.
- Reads the XML and looks for patterns like:
  - `$<` (a `$` followed by an XML tag opening)
  - `${` that later closes with `}`
- For each detected fragment, it:
  - Collects everything from `${` to `}` even if it crosses tags.
  - Removes XML tags and unwanted characters inside that range.
  - Replaces the original broken substring (for example `${NA    ME}`) with a clean placeholder `${NAME}`.
- Writes the fixed XML to `document.fix.xml`.
- Puts it back into the DOCX (replacing `word/document.xml`).
- Cleans up the temporary folder.

You can also see detailed logs in the browser if you tick **Show Debug Only** or **Show All Template Tag**.

---

## Limitations

- Designed specifically for PHPWord-style placeholders `${...}`.
- Focused on the common case where Word splits placeholders across runs or inserts tabs/whitespace.
- Not a generic DOCX/XML fixer – it only touches text around `${...}` patterns.

---

