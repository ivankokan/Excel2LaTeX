Attribute VB_Name = "Tools"
Option Explicit

Public Function Printf(ByVal sFormat As String, ParamArray Values()) As String
    Dim lValuesUBound As Long
    Dim sResult As String
    Dim vElement As Variant
    Dim lText As Long
    Dim aText() As String
    Dim sFirstChar As String
    Dim lValuePos As Long
    Dim sCurrentValue As String
    
    If IsMissing(Values()) Then
        lValuesUBound = -1
    Else
        lValuesUBound = UBound(Values)
    End If
    
    ' Handle all tokens:
    aText = Split(sFormat, "%")
    
    ' First entry of aText is text until the first occurence of %
    ' Start from second entry:
    For lText = LBound(aText) + 1 To UBound(aText)
        sFirstChar = Left$(aText(lText), 1)
        Select Case sFirstChar
        Case "1" To "9"
            ' Positional parameter: Lookup and insert
            lValuePos = CLng(sFirstChar) - 1
            
            If lValuePos <= lValuesUBound Then
                sCurrentValue = Values(lValuePos)
            Else
                ' Default: E.g., keep %3 if only two parameters are passed
                sCurrentValue = "%" & sFirstChar
            End If
            aText(lText) = sCurrentValue & Mid$(aText(lText), 2)
        
        Case "%"
            Debug.Assert False
            
        Case ""
            ' Special case: %% (or % at end of string):
            ' keep single % and ignore next token
            aText(lText) = "%" & aText(lText)
            lText = lText + 1
            
        Case Else
            ' Silently ignore all other %x tokens
            aText(lText) = "%" & aText(lText)
        End Select
    Next
    
    ' Combine result:
    Printf = Join(aText, "")
End Function

Public Sub SplitKeyValue(ByVal sKeyValue As String, ByRef sKey As String, ByRef sValue As String)
    Dim lPos As Long
    lPos = VBA.InStr(1, sKeyValue, "=")
    
    If lPos <= 0 Then
        sKey = sKeyValue
        sValue = ""
    Else
        sKey = Left$(sKeyValue, lPos - 1)
        sValue = Mid$(sKeyValue, lPos + 1)
    End If
End Sub















Private Sub Test_Printf()
    Debug.Assert Printf("%1", "abc") = "abc"
    Debug.Assert Printf("This is a %2%1.", "test", "(not too simple) ") = "This is a (not too simple) test."
    Debug.Assert Printf("Let's see how it handles out-of-range parameters %3 and occurences of %% and %y, and even at end: %") = "Let's see how it handles out-of-range parameters %3 and occurences of % and %y, and even at end: %"
    Debug.Assert Printf("%1%%%2%%%") = "%1%%2%%"
End Sub

Private Sub Test_SplitKeyValue()
    Dim sKey As String
    Dim sValue As String
    
    SplitKeyValue "ab=cd", sKey, sValue
    Debug.Assert sKey = "ab"
    Debug.Assert sValue = "cd"
    
    SplitKeyValue "ab=cd=ef", sKey, sValue
    Debug.Assert sKey = "ab"
    Debug.Assert sValue = "cd=ef"
    
    SplitKeyValue "abc", sKey, sValue
    Debug.Assert sKey = "abc"
    Debug.Assert sValue = ""
End Sub
