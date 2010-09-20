Attribute VB_Name = "Conversion"
' Excel2LaTeX:  is en excel to Latex converter.
' The improvements of V2.0 are based on Modifications by German Riano german@mendozas.com
' Changes introduced:
' * Graphical user interface
' * The LATeX code can be copied to clipboard and then pasted into you editor.
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
Public FullText

Sub Latex()
Attribute Latex.VB_Description = "Converts the selection to LaTex"
Attribute Latex.VB_ProcData.VB_Invoke_Func = "l\n14"
Dim name, selRange As Range
  If Selection Is Nothing Then GoTo ErrorMsg
  If TypeName(Selection) <> "Range" Then GoTo ErrorMsg
  If Selection.Areas.Count > 1 Then GoTo ErrorMsg
  
  Load frmConvert
  Set selRange = Selection
On Error GoTo NoName
  name = selRange.name.name
  GoTo continue
NoName:
  name = ActiveSheet.name
continue:
  On Error GoTo 0
  frmConvert.txtFilename = CurDir + "\" + name + ".tex"
  frmConvert.ConvertSelection
  frmConvert.Show
  Exit Sub
ErrorMsg:
  MsgBox "This macro coverts the selected table to Latex. Pleas select a single table", vbOKOnly + vbCritical
End Sub

Function NewModel() As Model
    Set NewModel = New Model
End Function
