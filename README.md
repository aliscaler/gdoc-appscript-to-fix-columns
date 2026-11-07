# gdoc-appscript-to-fix-columns

A small Google Apps Script for normalizing table formatting in Google Docs.

This script fixes common layout issues in document tables by:
- setting uniform cell padding
- removing empty trailing paragraphs inside cells
- resetting spacing and line height for paragraphs
- reducing hidden empty content that causes table rows to look uneven or misaligned

It is especially useful when copied content from other sources creates inconsistent table formatting.

## What it does

The script loops through every table in the active Google Doc and applies a cleaner, consistent layout:

- sets cell padding to 0.1 inch (matching the Google Docs UI default)
- removes extra empty paragraphs at the end of cells
- resets paragraph spacing before/after to 0
- sets line spacing to 1.0
- hides empty paragraphs by shrinking their font size to 1
- saves and closes the document after processing

## How to use

1. Open your Google Document.
2. Go to Extensions → Apps Script.
3. Paste the following script into the script editor.
4. Save the project.
5. Run the function `makeAllTablesExact01()`.
6. Allow the script to access the document if prompted.

```javascript
function makeAllTablesExact01() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const tables = body.getTables();

  if (tables.length === 0) {
    Logger.log("No tables found in this document.");
    return;
  }

  tables.forEach(table => {
    const numRows = table.getNumRows();
    for (let r = 0; r < numRows; r++) {
      const row = table.getRow(r);
      row.setMinimumHeight(0);

      const numCells = row.getNumCells();
      for (let c = 0; c < numCells; c++) {
        const cell = row.getCell(c);

        // 0.1 inches = 7.2 points (this matches the UI 0.1 in setting)
        cell.setPaddingTop(7.2);
        cell.setPaddingBottom(7.2);
        cell.setPaddingLeft(7.2);
        cell.setPaddingRight(7.2);

        // Clean up empty trailing paragraphs
        let numParagraphs = cell.getNumChildren();
        for (let p = numParagraphs - 1; p >= 1; p--) {
          const child = cell.getChild(p);
          if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
            const paragraph = child.asParagraph();
            if (paragraph.getText().trim() === "") {
              paragraph.removeFromParent();
            }
          }
        }

        // Strip line spacing and space before/after from all remaining paragraphs
        for (let p = 0; p < cell.getNumChildren(); p++) {
          const child = cell.getChild(p);
          if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
            const paragraph = child.asParagraph();
            paragraph.setSpacingBefore(0);
            paragraph.setSpacingAfter(0);
            paragraph.setLineSpacing(1.0);

            if (paragraph.getText().trim() === "") {
              paragraph.setFontSize(1);
            }
          }
        }
      }
    }
  });

  doc.saveAndClose();
}
