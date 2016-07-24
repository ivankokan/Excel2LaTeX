<h1>Excel2LaTeX</h1>
Making tables in LaTeX can be tedious, especially if some columns are calculated.
This converter allows you to write a table in Excel instead, and export the current selection as LaTeX markup
which can be pasted into an existing LaTeX document, or exported to a file and included via the `\input` command.

Known to be compatible with Windows Excel 2000&ndash;2016 (32-bit and 64-bit) and Mac Excel 2004, 2011, and 2016.
May also be compatible with other versions that support `.xla` add-ins.

![Excel and Excel2LaTeX comparison](https://i.imgur.com/UNKCihT.png)

<h2>Features</h2>
Most Excel formatting is supported.
 * Bold and italic (if applied to the whole cell)
 * Left, right, center, and general alignment (per-cell or per-column)
 * Vertical and horizontal borders (per-cell or per-column, single or double)
 * Font color (using the `xcolor` package)
 * Fill color (using the `colortbl` package)
 * Rotation (using the `rotating` package)
 * Merged cells (using the `multirow` package, if needed)
 * Can convert `\`, `$`, `_`, `^`, `%`, `&`, and `#` to appropriate macros, or leave them in-place
 * Supports `booktabs` package
 * Uses `bigstrut` package when `booktabs` is not available
 * Makes standard LaTeX `tabular` environment
 * Can surround `tabular` environment with `table` environment template
 * Copy output to clipboard or export to a `.tex` file for inclusion using `\include`
 * Save table specifications to your Excel worksheet, then export all tables at once


<h2>Using</h2>
Just open the file Excel2LaTeX.xla in Excel.  Then you will have two additional 
menu items in your **Tools** menu and a new toolbar with two buttons on it.  For 
Excel 2007 and later, you will have two new buttons in the **Add-Ins** ribbon.  If 
you plan to use the program frequently, you can save it in your addin directory 
and add it with **Tools**→**Add-Ins**.  This way it will be loaded whenever Excel is 
opened.

Select the table to convert and hit the button **Convert Table to LaTeX**.  You 
will be given the option to save the result to a `.tex` file, or send it to the clipboard 
(so you can paste it into your LaTeX editor).  Hit the **Store** button to store the 
current settings so you can **Load** them later or **Export All** to files.

![Excel2LaTeX interface](https://i.imgur.com/EK88upo.png)

<h2>Contributing</h2>
To work with the project, you will require my [VBA Sync Tool](https://github.com/chelh/VBASync).
