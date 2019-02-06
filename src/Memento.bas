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
    ElseIf pModel.Encoding = MsoEncoding.msoEncodingUTF8 Then
        ' Before writing in Binary mode the file must be opened and closed in Output mode
        ' (to erase all potential previous content)
        Open sFileName For Output Access Write Lock Read Write As #1
        Close #1
        
        Open sFileName For Binary Access Write Lock Read Write As #1
        Dim aUnicodes() As Long
        aUnicodes = StringToUnicodes(sCodePageRemark & vbNewLine & pModel.GetConversionResult)
        Dim l As Long
        For l = LBound(aUnicodes) To UBound(aUnicodes)
            Put #1, , UnicodeToUtf8Bytes(aUnicodes(l))
        Next
        Close #1
    End If
End Sub


Public Function UnicodeToUtf8Bytes(ByVal lUnicode As Long) As Byte()
    Const UNICODE_CODEPOINT_MIN As Long = &H0&
    Const UNICODE_CODEPOINT_MAX1 As Long = &H7F&
    Const UNICODE_CODEPOINT_MAX2 As Long = &H7FF&
    Const UNICODE_CODEPOINT_MAX3 As Long = &HFFFF&
    Const UNICODE_CODEPOINT_MAX4 As Long = &H10FFFF
    
    Dim aUtf8Bytes() As Byte
    
    Debug.Assert UNICODE_CODEPOINT_MIN <= lUnicode And lUnicode <= UNICODE_CODEPOINT_MAX4
    Select Case lUnicode
        Case UNICODE_CODEPOINT_MIN To UNICODE_CODEPOINT_MAX1
            ReDim aUtf8Bytes(1 To 1) As Byte
            aUtf8Bytes(1) = &H0& Or (lUnicode)
        
        Case UNICODE_CODEPOINT_MAX1 + 1 To UNICODE_CODEPOINT_MAX2
            ReDim aUtf8Bytes(1 To 2) As Byte
            aUtf8Bytes(2) = &H80& Or (lUnicode Mod &H40&)
            lUnicode = lUnicode \ &H40&
            aUtf8Bytes(1) = &HC0& Or (lUnicode)
        
        Case UNICODE_CODEPOINT_MAX2 + 1 To UNICODE_CODEPOINT_MAX3
            ReDim aUtf8Bytes(1 To 3) As Byte
            aUtf8Bytes(3) = &H80& Or (lUnicode Mod &H40&)
            lUnicode = lUnicode \ &H40&
            aUtf8Bytes(2) = &H80& Or (lUnicode Mod &H40&)
            lUnicode = lUnicode \ &H40&
            aUtf8Bytes(1) = &HE0& Or (lUnicode)
        
        Case UNICODE_CODEPOINT_MAX3 + 1 To UNICODE_CODEPOINT_MAX4
            ReDim aUtf8Bytes(1 To 4) As Byte
            aUtf8Bytes(4) = &H80& Or (lUnicode Mod &H40&)
            lUnicode = lUnicode \ &H40&
            aUtf8Bytes(3) = &H80& Or (lUnicode Mod &H40&)
            lUnicode = lUnicode \ &H40&
            aUtf8Bytes(2) = &H80& Or (lUnicode Mod &H40&)
            lUnicode = lUnicode \ &H40&
            aUtf8Bytes(1) = &HF0& Or (lUnicode)
    End Select
    
    UnicodeToUtf8Bytes = aUtf8Bytes
End Function


' Strings in VBA are encoded in UTF-16LE (http://www.di-mgt.com.au/howto-convert-vba-unicode-to-utf8.html)
' The largest UTF-16 code point is U+10FFFF (decimal 1114111), hence Long() as return type
Public Function StringToUnicodes(ByVal s As String) As Long()
    Const UTF16_CODEPOINT_MIN As Long = &H0&
    Const UTF16_SURROGATE_MIN As Long = &HD800&
    Const UTF16_SURROGATE_MAX As Long = &HDFFF&
    Const UTF16_CODEPOINT_MAX As Long = &H10FFFF
    
    Dim aUnicodes() As Long
    ReDim aUnicodes(1 To Len(s)) As Long
    Dim lU As Long
    
    Dim aBytes() As Byte
    aBytes = s
    
    Dim lUnicode As Long, lFirst As Long, lSecond As Long
    Dim l As Long
    For l = LBound(aBytes) To UBound(aBytes) Step 2
        lFirst = aBytes(l) + aBytes(l + 1) * &H100&
        Debug.Assert UTF16_CODEPOINT_MIN <= lFirst And lFirst <= UTF16_CODEPOINT_MAX
        
        If UTF16_SURROGATE_MIN <= lFirst And lFirst <= UTF16_SURROGATE_MAX Then
            ' lFirst represents high (leading) surrogate
            ' And two more bytes representing low (trailing) surrogate must be processed
            l = l + 2
            lSecond = aBytes(l) + aBytes(l + 1) * &H100&
            Debug.Assert UTF16_CODEPOINT_MIN <= lSecond And lSecond <= UTF16_CODEPOINT_MAX
            
            lUnicode = (lFirst - &HD800&) * &H400& + (lSecond - &HDC00&) + &H10000
        Else
            lUnicode = lFirst
        End If
        
        lU = lU + 1
        aUnicodes(lU) = lUnicode
    Next
    
    ReDim Preserve aUnicodes(1 To lU) As Long
    StringToUnicodes = aUnicodes
End Function


Public Sub SaveAllStoredItems(ByVal pStorage As IStorage)
    Dim cItems As Collection
    Set cItems = pStorage.GetItems
    
    Dim l1 As Long
    For l1 = 1 To cItems.Count
        SaveConversionResultToFile cItems(l1)
    Next
End Sub
