Attribute VB_Name = "Conversion"
' Excel2LaTeX: The Excel add-in for creating LaTeX tables
'
' Converts the selected cells to a LaTeX table, that can be included in a .tex file
' via \input{"table.tex"} or that can be copied to the clipboard. Most of the formatting
' is converted too. You can put additional LaTeX code in the cells, which will remain
' untouched by the converter.
'
' Copyright (c) 1996–2016 Chelsea Hughes, Kirill Müller, Andrew Hawryluk, Germán Riaño,
' and Joachim Marder.
'
' This work is distributed under the LaTeX Project Public License, version 1.3 or later,
' available at http://www.latex-project.org/lppl.txt
'
' Chelsea Hughes currently maintains this project (comprising Excel2LaTeX.xla and
' README.md) and will receive error reports at the project GitHub page,
' https://github.com/krlmlr/Excel2LaTeX

Option Explicit

Public Sub LaTeX()
Attribute LaTeX.VB_Description = "Opens the main dialog for converting into LaTeX"
Attribute LaTeX.VB_ProcData.VB_Invoke_Func = "l\n14"
    With NewController
        Set .View = NewView
        Set .Model = NewDefaultModel
        Set .Storage = NewStorage
        .Run
    End With
End Sub

Public Sub LaTeXAllToFiles()
Attribute LaTeXAllToFiles.VB_Description = "Converts all configured selections into LaTeX"
Attribute LaTeXAllToFiles.VB_ProcData.VB_Invoke_Func = "l\n14"
    SaveAllStoredItems NewStorage
End Sub
