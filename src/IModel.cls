VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IModel"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public Enum x2lOptions
    x2lNone = &H0
    
    x2lCreateTableEnvironment = &H1
    x2lBooktabs = &H2
    x2lConvertMathChars = &H4
End Enum

Public Property Get Events() As IModelEvents
'
End Property

Public Property Get Options() As x2lOptions
'
End Property
Public Property Let Options(ByVal eOptions As x2lOptions)
'
End Property

Public Property Get Indent() As Integer
'
End Property
Public Property Let Indent(ByVal iIndent As Integer)
'
End Property

Public Property Get CellWidth() As Integer
'
End Property
Public Property Let CellWidth(ByVal iCellWidth As Integer)
'
End Property

Public Property Get FileName() As String
'
End Property
Public Property Let FileName(ByVal sFileName As String)
'
End Property
Public Property Get AbsoluteFileName() As String
'
End Property

Public Property Get RangeAddress() As String
'
End Property
Public Property Let RangeAddress(ByVal sRangeAddress As String)
'
End Property

Public Function GetConversionResult() As String
'
End Function

Public Sub InitDefault()
'
End Sub

Public Property Get Description() As String
'
End Property

Public Property Get Range() As Range
'
End Property
Public Property Set Range(ByVal rRange As Range)
'
End Property
