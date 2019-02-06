Attribute VB_Name = "Memento"
Option Explicit

Public aEncodings() As Variant

Public Function ModelPropertyNames() As String()
    Const NAMES As String = "RangeAddress|Options|CellWidth|Indent|FileName|Encoding"
    ModelPropertyNames = Split(NAMES, "|")
End Function

Public Function ModelToCollection(ByVal pModel As IModel) As Collection
    Set ModelToCollection = New Collection
    
    Dim sName As Variant
    For Each sName In ModelPropertyNames()
        ModelToCollection.Add CallByName(pModel, sName, VbGet), sName
    Next
End Function

Public Function ModelToString(ByVal pModel As IModel) As String
    Dim sName As Variant
    For Each sName In ModelPropertyNames()
        ModelToString = ModelToString & Printf("%1=%2;", sName, CallByName(pModel, sName, VbGet))
    Next
End Function

Public Sub CollectionToModel(ByVal pModel As IModel, ByVal cCollection As Collection)
    Dim sName As Variant
    For Each sName In ModelPropertyNames()
        On Error Resume Next
        CallByName pModel, sName, VbLet, cCollection(sName)
        On Error GoTo 0
    Next
End Sub

Public Sub StringToModel(ByVal pModel As IModel, ByVal sSettings As String)
    Dim aSettings() As String
    aSettings = Split(sSettings, ";")
    
    Dim l1 As Long
    Dim sKey As String
    Dim sValue As String
    For l1 = 0 To UBound(aSettings)
        SplitKeyValue aSettings(l1), sKey, sValue
        
        On Error Resume Next
        CallByName pModel, sKey, VbLet, sValue
        On Error GoTo 0
    Next
End Sub

Public Function CollectionToNewModel(ByVal cSettings As Collection) As IModel
    Set CollectionToNewModel = NewModel()
    CollectionToModel CollectionToNewModel, cSettings
End Function

Public Function StringToNewModel(ByVal sSettings As String) As IModel
    Set StringToNewModel = NewModel()
    StringToModel StringToNewModel, sSettings
End Function


Public Function RangeToAddress(ByVal rRange As Range) As String
    If rRange Is Nothing Then Exit Function
    RangeToAddress = Printf("'%1'!%2", rRange.Worksheet.Name, rRange.Address)
End Function

Public Function AddressToRange(ByVal sRangeAddress As String) As Range
    Set AddressToRange = Nothing
    If sRangeAddress = "" Then Exit Function
    Set AddressToRange = Application.Range(sRangeAddress)
End Function


Private Function GetEncoding(ByVal aEncodings As Variant, ByVal eEncoding As MsoEncoding) As Variant
    Dim aEncoding As Variant
    For Each aEncoding In aEncodings
        If aEncoding(0) = eEncoding Then
            GetEncoding = aEncoding
            Exit Function
        End If
    Next
End Function

Public Sub SaveConversionResultToFile(ByVal pModel As IModel)
    Dim sFileName As String
    sFileName = pModel.AbsoluteFileName
    If sFileName = "" Then Exit Sub
    
    Dim sCodePageRemark As String, aEncoding As Variant
    aEncoding = GetEncoding(aEncodings, pModel.Encoding)
    Debug.Assert Not IsEmpty(aEncoding)
    sCodePageRemark = Printf("% codepage %1 (%2)", aEncoding(0), aEncoding(1))
    
    If pModel.Encoding = Application.DefaultWebOptions.Encoding Then
        Open sFileName For Output Access Write Lock Read Write As #1
        Print #1, sCodePageRemark
        Print #1, pModel.GetConversionResult;
        Close #1
    Else
        ' Before writing in Binary mode the file must be opened and closed in Output mode
        ' (to erase all potential previous content)
        Open sFileName For Output Access Write Lock Read Write As #1
        Close #1
        
        Dim se As StringEncoder
        Set se = New StringEncoder
        
        Dim aBytes() As Byte
        aBytes = se.Encode(sCodePageRemark & vbNewLine & pModel.GetConversionResult, pModel.Encoding)
        Open sFileName For Binary Access Write Lock Read Write As #1
        Put #1, , aBytes
        Close #1
    End If
End Sub

Public Sub SaveAllStoredItems(ByVal pStorage As IStorage)
    Dim cItems As Collection
    Set cItems = pStorage.GetItems
    
    Dim l1 As Long
    For l1 = 1 To cItems.Count
        SaveConversionResultToFile cItems(l1)
    Next
End Sub
