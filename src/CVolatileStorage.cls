VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVolatileStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IStorage

Private mCollection As New Collection


Private Function IStorage_Add(ByVal pModel As IModel) As Long
    IStorage_Add = mCollection.Count + 1
    mCollection.Add pModel
End Function

Private Sub IStorage_Remove(ByVal lIndex As Long)
    mCollection.Remove lIndex
End Sub

Private Function IStorage_GetItems() As Collection
    Set IStorage_GetItems = New Collection
    
    Dim pModel As IModel
    For Each pModel In mCollection
        IStorage_GetItems.Add pModel
    Next
End Function

