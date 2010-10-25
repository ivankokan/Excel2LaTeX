Attribute VB_Name = "Memento"
Option Explicit

Public Function ModelToString(ByVal pModel As IModel) As String
    With pModel
        ModelToString = "" _
            & Printf("Options=%1;", .Options) _
            & Printf("CellWidth=%1;", .CellWidth) _
            & Printf("Indent=%1;", .Indent) _
            & Printf("RangeAddress=%1;", .RangeAddress) _
            & Printf("FileName=%1;", .FileName) _
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
            Case "RangeAddress"
                .RangeAddress = sValue
            Case "FileName"
                .FileName = sValue
            End Select
        End With
        On Error GoTo 0
    Next
End Sub

Public Function StringToNewModel(ByVal sSettings As String) As IModel
    Set StringToNewModel = NewModel()
    StringToModel StringToNewModel, sSettings
End Function


Public Function RangeToAddress(ByVal pRange As Range) As String
    If pRange Is Nothing Then Exit Function
    RangeToAddress = Printf("%1!%2", pRange.Worksheet.Name, pRange.Address)
End Function

Public Function AddressToRange(ByVal sRangeAddress As String) As Range
    Set AddressToRange = Nothing
    If sRangeAddress = "" Then Exit Function
    Set AddressToRange = Application.Range(sRangeAddress)
End Function

Public Sub SaveConversionResultToFile(ByVal pModel As IModel)
    Dim sFileName As String
    sFileName = pModel.AbsoluteFileName
    If sFileName = "" Then Exit Sub
    
    Open sFileName For Output As 1
    Print #1, pModel.GetConversionResult
    Close #1
End Sub
