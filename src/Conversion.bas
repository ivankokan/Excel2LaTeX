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
    NewController.LaTeX NewView, NewDefaultModel
End Sub

Function NewController() As IController
    Set NewController = New CController
End Function

Private Function NewModel() As IModel
    Set NewModel = New CModel
End Function

Private Function NewDefaultModel() As IModel
    Set NewDefaultModel = NewModel
    NewDefaultModel.InitDefault
End Function

Function NewView() As frmConvert
    Set NewView = New frmConvert
End Function

Public Function ModelToString(ByVal pModel As IModel) As String
    With pModel
        ModelToString = "" _
            & Printf("Options=%1;", .Options) _
            & Printf("CellWidth=%1;", .CellWidth) _
            & Printf("Indent=%1;", .Indent) _
            & ""
    End With
End Function

Public Sub StringToModel(ByVal pModel As IModel, ByVal sSettings As String)
    Dim aSettings() As String
    aSettings = Split(sSettings, ";")
    
    Dim l1 As Long
    Dim sKey As String
    Dim sValue As String
    For l1 = 0 To UBound(aSettings)
        SplitKeyValue aSettings(l1), sKey, sValue
        
        On Error Resume Next
        With pModel
            Select Case sKey
            Case "Options"
                .Options = sValue
            Case "CellWidth"
                .CellWidth = sValue
            Case "Indent"
                .Indent = sValue
            End Select
        End With
        On Error GoTo 0
    Next
End Sub
