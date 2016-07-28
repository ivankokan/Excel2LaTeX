VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Const INITIAL_SIZE = 16&

Private m_buf As String 'current buffer
Private m_len As Long 'size in characters

Private Sub Class_Initialize()
    m_buf = Space$(INITIAL_SIZE)
End Sub

Public Property Get Length() As Long
    Length = m_len
End Property

Public Function Append(s As String) As StringBuilder
    If m_len + Len(s) > 2 * Len(m_buf) Then
        m_buf = m_buf & Space$(2 * Len(s))
    ElseIf m_len + Len(s) > Len(m_buf) Then
        m_buf = m_buf & Space$(Len(m_buf))
    End If
    If Len(s) > 0 Then
        Mid(m_buf, m_len + 1, Len(s)) = s
        m_len = m_len + Len(s)
    End If
    
    Set Append = Me
End Function

Public Function ToString() As String
    ToString = Left$(m_buf, m_len)
End Function
