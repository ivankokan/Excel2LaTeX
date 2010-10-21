Attribute VB_Name = "Conversion"
' Excel2LaTeX:  is en excel to LaTeX converter.
' The improvements of V2.0 are based on Modifications by German Riano german@mendozas.com
' Changes introduced:
' * Graphical user interface
' * The LaTeX code can be copied to clipboard and then pasted into you editor.
' * Better handling of multicolum cells
' * doublelines on top border are now handled
'
' Converts the selected cells to a LaTeX table, that can be included in a tex-file
' via \input{"table.tex"} ot that can be copied to the clipboard. Most of the formatting
' is converted too. You can put additional LaTeX code in the cells, which will remain
' untouched by the converter.
'
' This converter is freeware. You can freely use and distribute it
' © 1996-2001 by Joachim Marder and German Riano
'
'
' Send bug reports and suggestions to: marder@jam-software.com
' Web Page for Excel2LaTeX: http://www.jam-software.de/software.html
'

Option Explicit

Sub LaTeX()
Attribute LaTeX.VB_Description = "Converts the selection to LaTex"
Attribute LaTeX.VB_ProcData.VB_Invoke_Func = "l\n14"
    NewController.LaTeX NewView, NewDefaultModel, NewStorage
End Sub

