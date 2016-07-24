**Version 3.4.0: Not currently released**
 * Bug fix: Wasn't loading properly on Excel Mac 2016
 * Bug fix: Changing the selected range after closing the form would sometimes cause an error
 * Changed form layout and added a checkbox which arrests updates when unticked
 * Booktabs mode won't force three-line table format
 * Support conditional/table formatting in Excel 2010+
 * Color is now supported in fonts and fills
 * If a single cell is selected, the form will attempt to convert the entire region
 * Conversion now respects left-alignment of text in "General" mode
 * Special characters will always be converted if a cell is numeric
 * Pared down the About form

**Version 3.3: Released on Sep 27 2012**
 * Bug fix: Doesn't crash when trying to start conversion when an entire line is selected.
 * Bug fix: Better conversion of backslashes: `\textbackslash{}` instead of `\textbackslash `
 * Performance: Improvements by avoiding unnecessary calls to Excel objects

**Version 3.2: Released on Mar 26 2012**
 * Bug fix: Finally restored compatibility with Office Mac
 * Bug fix: Do not add extra alignment tab after `\multicolumn{}{}`

**Version 3.1: Released on May 19 2011**
 * If the column width is set to 0, each cell occupies a separate line in the output file
 * In booktabs mode, no vertical space is inserted before the top row anymore
 * Bug fix: Restored compatibility with Office 2000 and Office Mac
 * Bug fix: Form is protected against erroneous entries

**Version 3.0: Released on Nov 17 2010**
 * **CAVEAT:** The toolbar buttons and menu items from previous versions of Excel2LaTeX are not deleted automatically.
 * **CAVEAT:** Depending on the formatting of your table, you might require the following packages not required before:
   * `bigstrut`
   * `rotating`
   * `multirow`
 * Stored tables: See annotations above.
 * More cell formatting options are used in the converted table.
 * Cells with rotated contents are supported. Requires the `rotating` package.
 * Multirow and multirow-multicolumn cells are supported.  Requires the `multirow` package.
 * A cell formatted as bold and containing inline math formulae is typeset in bold using `\boldmath` and `\unboldmath`
 * More precise typesetting of cell borders in non-booktabs mode.
   * A column is assumed to have a vertical border if there is a border in any row of this column. Cells with a vertical
     border different from the default or without vertical border are typeset using the `\multicol` command, specifying 
     the border type for this cell.  If both single and double lines are present in one column, a double line is assumed
     as default for this column.
   * Horizontal rules are typeset using `\cline` if they do not go straight through from left to right.  Double horizontal
     lines are converted to single lines in this case.
   * `\bigstrut` commands are inserted where appropriate.  The number of struts required for a multirow cell is computed
     correctly.  Requires the `bigstrut` package.
 * Bug fix: The type of the left border of a multicolumn cell is determined correctly (in non-booktabs mode).
 * Improved file name handling.
   * The target `.tex` file will be stored in the directory of the Excel worksheet by default.
   * If the target `.tex` file resides in the directory of the worksheet or in a subdirectory, a relative path to the file is stored.
 * The main form is now modeless.  The worksheet can be edited while the form is open.
   * Changes to the contents of the selection are tracked, the LaTeX table in the text box is updated automatically.
     Changes to cell formatting (font style, borders) are not tracked.
   * The current selection can be set as source range for the current conversion to LaTeX by hitting the large button at the top
     of the dialog.
 * The main form shows up always, even if no range or a multi-area range is selected.

**Version 2.3: Released on Nov 16 2010**
 * Bug fix: In Office 2007, no error is raised after opening a document anymore.
 * Bug fix: When writing the TeX file, no additional newline is appended. Spurious spaces may produce unwanted results.

**Version 2.2: Released on Sep 29 2010**
 * Save and load settings to/from registry.
 * Bug fix: do not add two command buttons to the ribbons in Excel 2007 and later.
 * Bug fix: use `\textbf` and `\textit` instead of `\bf` and `\it`.
 * Bug fix: do not use vertical borders for multicolumn environment for booktabs tables.
 * Bug fix: correctly determine LaTeX column borders if Excel cell borders are set only for the top and/or left border.
 * Internal: avoid copying the range to a hidden worksheet before converting it to LaTeX.
 * Internal: various code refactorings.
  
**Version 2.1: Released on Sep 18 2008**
 * Better character replacement: the previous version only replaced the first occurrence of $ or % in a cell.
 * Optionally generate a table environment, format the table in the style of the `booktabs` LaTeX package, and/or add
   extra leading indent spaces.
 * Better interactivtiy (no refresh button required).
 * Bug fix: the previous version would damage formulas that referred to cells outside the selection.
  
**Version 2.0: Released on Jul 21 2001**
This version is based on modifications by Germán Riaño
 * Graphical user interface
 * The LaTeX code can be copied to clipboard and then pasted into you editor.
 * Better handling of multicolum cells
 * doublelines on top border are now handled

**Version 1.2: Published on Nov 22 1998**
 * The characters % and $ are now converted to the correspondig LaTeX macros

**Version 1.1: Published on Apr 12 1997**
 * Some small changes to make it run with Excel 97 too

**Version 1.0: Published on Oct 22 1996**
 * Initial release
