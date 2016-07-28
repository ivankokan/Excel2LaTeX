Attribute VB_Name = "Win32Clipboard"
Option Explicit

' this code adapted from Stack Overflow: "Excel 2013 64-bit VBA: Clipboard API doesn't work"
' http://stackoverflow.com/q/18668928/2146688

' changed to use CF_UNICODETEXT, and RtlMoveMemory instead of lstrcpy

#If VBA7 And Win32 Then
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal DestPtr As LongPtr, ByVal SrcPtr As LongPtr, ByVal sz As Long)

Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

Public Const CF_UNICODETEXT = 13
Public Const CB_MAXSIZE = 4096

Public Function Win32_SetClipBoard(MyString As String) As Boolean
'32-bit code by Microsoft: http://msdn.microsoft.com/en-us/library/office/ff192913.aspx
    Dim hGlobalMemory As LongPtr, lpGlobalMemory As LongPtr
    Dim hClipMemory As LongPtr, X As Long

    ' Allocate moveable global memory.
    hGlobalMemory = GlobalAlloc(GHND, Len(MyString) * 2 + 2)
    If hGlobalMemory = 0 Then MsgBox "Could not allocate memory.": Exit Function

    ' Lock the block to get a far pointer to this memory.
    lpGlobalMemory = GlobalLock(hGlobalMemory)

    ' Copy the string to this global memory.
    CopyMemory lpGlobalMemory, StrPtr(MyString), Len(MyString) * 2

    ' Unlock the memory.
    If GlobalUnlock(hGlobalMemory) <> 0 Then
       MsgBox "Could not unlock memory location. Copy aborted."
       'Debug.Print "GlobalFree returned: " & CStr(GlobalFree(hGlobalMemory))
       GoTo OutOfHere
    End If

    ' Open the Clipboard to copy data to.
    If OpenClipboard(0&) = 0 Then
       MsgBox "Could not open the Clipboard. Copy aborted."
       Exit Function
    End If

    ' Clear the Clipboard.
    EmptyClipboard

    ' Copy the data to the Clipboard.
    SetClipboardData CF_UNICODETEXT, hGlobalMemory

OutOfHere:
    If CloseClipboard() = 0 Then
       MsgBox "Could not close Clipboard."
    End If
    Win32_SetClipBoard = True
End Function
#End If
