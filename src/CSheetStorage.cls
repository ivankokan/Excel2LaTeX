VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSheetStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IStorage

Private mpEvents As New IStorageEvents


Private Property Get IStorage_Events() As IStorageEvents
    Set IStorage_Events = mpEvents
End Property

Private Function IStorage_Add(ByVal pModel As IModel, Optional ByVal lIndex = -1) As Long
    mpEvents.RaiseChanged
End Function

Private Sub IStorage_Remove(ByVal lIndex As Long)
    mpEvents.RaiseChanged
End Sub

Private Function IStorage_GetItems() As Collection
    Set IStorage_GetItems = New Collection
End Function


