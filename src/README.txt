This is the "development environment". The master file is Excel2LaTeX.xls. 
The macros and forms in the master file are exported from these master files 
to plain text files to be able to track modifications, and also imported back 
to Excel2LaTeX.xla . That is, Excel2LaTeX.xla is built from scratch for every 
commit to the repository, and modifications to Excel2LaTeX.xla that are not 
backed up by corresponding modifications to src/Excel2LaTeX.xls are lost.

For enhancements to the Excel2LaTeX.xla, the following workflow is suggested:

- Edit the code in Excel2LaTeX.xls
- From the VBA debug window (accessible via Ctrl+G), call Diff to see the current changes
- To commit the changes, call Commit from the debug window. An instance of notepad.exe pops up, allowing you to enter a commit message
- Finally, push your changes to Launchpad

