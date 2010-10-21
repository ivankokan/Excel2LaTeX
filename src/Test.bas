Attribute VB_Name = "Test"
Option Explicit

Private Sub Test_Storage()
    Dim pStorage As IStorage
    Set pStorage = NewStorage
    
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
End Sub
