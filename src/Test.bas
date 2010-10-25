Attribute VB_Name = "Test"
Option Explicit

Private Sub Test_VolatileStorage()
    Test_Storage New CVolatileStorage
End Sub

Private Sub Test_SheetStorage()
    On Error Resume Next
    With ActiveWorkbook.Sheets("Excel2LaTeX")
        .Range.Clear
        .Delete
    End With
    On Error GoTo 0
    Test_Storage New CSheetStorage
End Sub

Private Sub Test_Storage(ByVal pStorage As IStorage)
    Dim lIndex As Long
    lIndex = pStorage.Add(NewDefaultModel())
    Debug.Assert lIndex = 1
    
    Dim pModel As IModel
    Set pModel = NewDefaultModel
    pModel.CellWidth = pModel.CellWidth + 1
    Dim lIndex2 As Long
    lIndex2 = pStorage.Add(pModel)
    Debug.Assert lIndex <> lIndex2
    
    pStorage.Remove lIndex
    
    Dim pCollection As Collection
    Set pCollection = pStorage.GetItems
    Debug.Assert pCollection.Count = 1
    Debug.Assert pCollection.Item(1).CellWidth = NewDefaultModel().CellWidth + 1
    
    pStorage.Add NewDefaultModel, 0
    pStorage.Add NewDefaultModel, pStorage.GetItems.Count
    Debug.Assert pStorage.GetItems.Count = 3
    
    pStorage.Remove 1
    pStorage.Remove 2
    pStorage.Remove 1
    Debug.Assert pStorage.GetItems.Count = 0
End Sub
