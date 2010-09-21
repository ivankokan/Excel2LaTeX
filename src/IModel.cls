VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum x2lOptions
    x2lNone = &H0
    
    x2lCreateTableEnvironment = &H1
    x2lBooktabs = &H2
    x2lConvertMathChars = &H4
End Enum

Public Property Get Options() As x2lOptions
'
End Property
Public Property Let Options(ByVal Options As x2lOptions)
'
End Property

Public Property Get Indent() As Integer
'
End Property
Public Property Let Indent(ByVal Indent As Integer)
'
End Property

Public Property Get CellWidth() As Integer
'
End Property
Public Property Let CellWidth(ByVal CellWidth As Integer)
'
End Property

Public Property Get FileName() As String
'
End Property
Public Property Let FileName(ByVal FileName As String)
'
End Property

