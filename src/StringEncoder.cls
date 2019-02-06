VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringEncoder"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit


Public Function Encode(ByVal s As String, ByVal eEncoding As MsoEncoding) As Byte()
    Dim aBytes() As Byte
    Dim b As Long
    
    Select Case eEncoding
        Case MsoEncoding.msoEncodingUTF8
            Dim aUnicodes() As Long
            aUnicodes = StringToUnicodes(s)
            ReDim aBytes(0 To 4 * (UBound(aUnicodes) - LBound(aUnicodes) + 1) - 1) As Byte
            
            Dim u As Long, aBytesTemp() As Byte, t As Long
            For u = LBound(aUnicodes) To UBound(aUnicodes)
                aBytesTemp = UnicodeToUtf8Bytes(aUnicodes(u))
                For t = LBound(aBytesTemp) To UBound(aBytesTemp)
                    aBytes(b) = aBytesTemp(t)
                    b = b + 1
                Next
            Next
            ReDim Preserve aBytes(0 To b - 1) As Byte
        
        Case MsoEncoding.msoEncodingUnicodeLittleEndian
            aBytes = s
        
        Case MsoEncoding.msoEncodingUnicodeBigEndian
            aBytes = s
            For b = LBound(aBytes) To UBound(aBytes) Step 2
                aBytes(b) = aBytes(b) Xor aBytes(b + 1)
                aBytes(b + 1) = aBytes(b) Xor aBytes(b + 1)
                aBytes(b) = aBytes(b) Xor aBytes(b + 1)
            Next
        
        Case Else
            MsgBox "Unsupported encoding!", vbCritical
    End Select
    
    Encode = aBytes
End Function


Private Function UnicodeToUtf8Bytes(ByVal lUnicode As Long) As Byte()
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
Private Function StringToUnicodes(ByVal s As String) As Long()
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
