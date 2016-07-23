VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_Open()
    CreateMenuItems
End Sub

Private Sub CreateMenuItem(ByVal Caption As String, ByVal Action As String, ByVal FaceID As Long)
    Dim myMenubar As CommandBar, toolsMenu As CommandBarPopup, newMenuItem As CommandBarControl
    Set myMenubar = Application.CommandBars("Worksheet Menu Bar")
    Set toolsMenu = myMenubar.Controls(6)
    Set newMenuItem = myMenubar.FindControl(Tag:=Action, Recursive:=True)
    If Not newMenuItem Is Nothing Then newMenuItem.Delete
    Set newMenuItem = toolsMenu.Controls.Add(Type:=msoControlButton, Before:=3, Temporary:=True)
    newMenuItem.Tag = Action
    newMenuItem.OnAction = Action
    newMenuItem.FaceID = FaceID
    newMenuItem.TooltipText = Caption
    newMenuItem.Caption = Caption
End Sub

Private Sub CreateMenuItems()
    #If Mac Then
        ' on Excel 2016 Mac, we can't manipulate menus
        If Val(Application.Version) >= 15 Then Exit Sub
    #End If
    CreateMenuItem "Con&vert Table to LaTeX", "Conversion.LaTeX", 8
    CreateMenuItem "Convert All Stored Tab&les to LaTeX", "Conversion.LaTeXAllToFiles", 107
    If Val(Application.Version) >= 12 Then
        ' don't show the toolbar on Excel 2007+, since it just shows up as duplicated buttons in the ribbon
        If FindCommandBar("Excel2LaTeX") > 0 Then Application.CommandBars("Excel2LaTeX").Delete
    End If
End Sub

Private Function FindCommandBar(ByVal Name As String)
    Dim i As Long
    For i = 1 To Application.CommandBars.Count
        If Application.CommandBars(i).Name = Name Then
            FindCommandBar = i
            Exit Function
        End If
    Next
End Function
