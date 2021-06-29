VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSheetStorage"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Implements IStorage

Private mpEvents As New IStorageEvents

Private Const MAX_DIM = 10000


Private Property Get IStorage_Events() As IStorageEvents
    Set IStorage_Events = mpEvents
End Property

Private Function IStorage_Add(ByVal pModel As IModel, Optional ByVal lIndex As Long = -1) As Long
    Dim lFreeRow As Long
    If lIndex < 0 Then
        lFreeRow = GetFreeRowIndex
    Else
        lFreeRow = lIndex + 2
        GetStorageSheet().Rows(lFreeRow).Insert xlShiftDown
    End If
    
    Dim rRow As Range
    Set rRow = GetStorageSheet().Rows(lFreeRow)
    
    Dim cColumns As Collection
    Set cColumns = GetColumnList()
    
    Dim cValues As Collection
    Set cValues = ModelToCollection(pModel)
    
    Dim sKey As Variant
    Dim lColumnIndex As Long
    For Each sKey In ModelPropertyNames()
        On Error Resume Next
        lColumnIndex = -1
        lColumnIndex = cColumns(sKey)
        On Error GoTo 0
        
        If lColumnIndex < 0 Then lColumnIndex = CreateColumn(sKey)
        Debug.Assert lColumnIndex > 0
        
        If sKey = "RangeAddress" Then
            AddressToFormula cValues(sKey), rRow.Cells(1, lColumnIndex)
        Else
            rRow.Cells(1, lColumnIndex) = cValues(sKey)
        End If
    Next
    
    IStorage_Add = lFreeRow - 1
    
    mpEvents.RaiseChanged
End Function

Private Sub IStorage_Remove(ByVal lIndex As Long)
    Dim rRow As Range
    Set rRow = GetStorageSheet().Rows(lIndex + 1)
    rRow.Delete xlShiftUp
    
    mpEvents.RaiseChanged
End Sub

Private Function IStorage_GetItems() As Collection
    Set IStorage_GetItems = New Collection
    
    Dim wStorageSheet As Worksheet
    Set wStorageSheet = GetStorageSheet(False)
    
    If wStorageSheet Is Nothing Then Exit Function

    Dim cColumns As Collection
    Dim cKeys As Collection
    Set cColumns = GetColumnList(cKeys)
    
    Dim cRow As Collection
    Dim lRow As Long
    Dim lColumn As Variant
    Dim sKey As String, sValue As String
    Dim rRow As Range
    For lRow = 2 To MAX_DIM
        Set rRow = wStorageSheet.Rows(lRow)
        If Application.WorksheetFunction.CountA(rRow) = 0 Then Exit For
        Set cRow = New Collection
        
        For Each lColumn In cColumns
            sKey = cKeys(lColumn)
            If sKey = "RangeAddress" Then
                sValue = FormulaToAddress(rRow.Cells(ColumnIndex:=lColumn))
            Else
                sValue = rRow.Cells(ColumnIndex:=lColumn)
            End If
            cRow.Add sValue, sKey
        Next
        
        IStorage_GetItems.Add CollectionToNewModel(cRow)
        Set cRow = Nothing
    Next
End Function


Private Sub AddressToFormula(ByVal sAddress As String, ByVal rCell As Range)
    rCell.Formula = Printf("=COUNT(%1)", sAddress)
End Sub

Private Function FormulaToAddress(ByVal rCell As Range) As String
    On Error Resume Next
    FormulaToAddress = Mid(rCell.Formula, 8, Len(rCell.Formula) - 8)
End Function


Private Function GetStorageSheet(Optional ByVal bCreate As Boolean = True) As Worksheet
    On Error Resume Next
    With ActiveWorkbook.Worksheets
        Set GetStorageSheet = .Item("Excel2LaTeX")
        If (Err.Number <> 0) And bCreate Then
            On Error GoTo 0
            Set GetStorageSheet = .Add
            GetStorageSheet.Name = "Excel2LaTeX"
            GetStorageSheet.Visible = xlSheetHidden
        End If
    End With
End Function

Private Function GetColumnList(Optional ByRef cNameList As Collection) As Collection
    Set GetColumnList = New Collection
    Set cNameList = New Collection
    
    Dim wStorageSheet As Worksheet
    Set wStorageSheet = GetStorageSheet(False)
    
    If wStorageSheet Is Nothing Then Exit Function
    
    Dim rFirstRow As Range
    Set rFirstRow = wStorageSheet.Rows(1)
    
    Dim l1 As Long
    Dim sName As String
    For l1 = 1 To MAX_DIM
        sName = rFirstRow.Cells(ColumnIndex:=l1)
        If sName = "" Then Exit For
        GetColumnList.Add l1, sName
        cNameList.Add sName
    Next
End Function

Private Function GetFreeRowIndex() As Long
    Dim wStorageSheet As Worksheet
    Set wStorageSheet = GetStorageSheet()
    
    Dim rFirstColumn As Range
    Set rFirstColumn = wStorageSheet.Columns(1)
    
    Dim l1 As Long
    Dim sValue As String
    For l1 = 2 To MAX_DIM
        sValue = rFirstColumn.Cells(RowIndex:=l1)
        If sValue = "" Then
            GetFreeRowIndex = l1
            Exit Function
        End If
    Next
End Function

Private Function GetFreeColumnIndex() As Long
    Dim wStorageSheet As Worksheet
    Set wStorageSheet = GetStorageSheet()
    
    Dim rFirstRow As Range
    Set rFirstRow = wStorageSheet.Rows(1)
    
    Dim l1 As Long
    Dim sValue As String
    For l1 = 1 To MAX_DIM
        sValue = rFirstRow.Cells(ColumnIndex:=l1)
        If sValue = "" Then
            GetFreeColumnIndex = l1
            Exit Function
        End If
    Next
End Function

Private Function CreateColumn(ByVal sName As String) As Long
    CreateColumn = GetFreeColumnIndex()
    
    Dim wStorageSheet As Worksheet
    Set wStorageSheet = GetStorageSheet()
    
    wStorageSheet.Cells(1, CreateColumn) = sName
End Function
