This is the "development environment". The master files are the .bas, .cls and 
.frm/.frx files. To edit these master files, the following workflow is suggested:

- After branch or checkout, open Excel2LaTeX.xls and execute the routine 
	Dev.CreateDevWorksheet. This imports the master files to a new Excel sheet,
	Excel2LaTeXDev.cls. This file is ignored for version control, modifications
	are not committed.
- From the VBA debug window (accessible via Ctrl+G), call Diff to see the current
	changes. This exports all code modules back to the master files and shows
	the differences within these files.
- To commit the changes, call Commit from the debug window. An instance of notepad.exe 
	pops up, allowing you to enter a commit message.
- If you change the Dev module, execute CreateBzrWorksheet before committing.
- If you want to publish a new version of the addin, execute ExportToAddin before 
	committing.
- Finally, push your changes to Launchpad.

Changes made directly to Excel2LaTeX.xla are lost eventually.
However, each call to Diff or Commit creates a development version of the addin
in the file Excel2LaTeXDev.xla, which is ignored for version control as well.
