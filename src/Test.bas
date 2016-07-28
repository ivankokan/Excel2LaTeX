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
    Debug.Assert pStorage.GetItems.Count = 1
    
    Debug.Assert pStorage.GetItems.Count = 1
    Debug.Assert pStorage.GetItems.Item(1).CellWidth = NewDefaultModel().CellWidth + 1
    
    pStorage.Add NewDefaultModel, 0
    Debug.Assert pStorage.GetItems.Item(2).CellWidth = NewDefaultModel().CellWidth + 1
    pStorage.Add NewDefaultModel, 2
    Debug.Assert pStorage.GetItems.Item(2).CellWidth = NewDefaultModel().CellWidth + 1
    pStorage.Add NewDefaultModel, 1
    Debug.Assert pStorage.GetItems.Item(3).CellWidth = NewDefaultModel().CellWidth + 1
    pStorage.Add NewDefaultModel, pStorage.GetItems.Count
    Debug.Assert pStorage.GetItems.Count = 5
    
    pStorage.Remove 1
    pStorage.Remove 2
    pStorage.Remove 3
    pStorage.Remove 2
    pStorage.Remove 1
    Debug.Assert pStorage.GetItems.Count = 0
End Sub

Private Sub Test_Model_AppendToRangeSet()
    Dim pModel As New CModel
    
    Dim sLineDef As String
    Dim lLineOpenFrom As Long
    
    pModel.AppendToRangeSet sLineDef, lLineOpenFrom, True, 1
    Debug.Assert sLineDef = "1"
    pModel.AppendToRangeSet sLineDef, lLineOpenFrom, False, 2
    Debug.Assert sLineDef = "1-1"
    pModel.AppendToRangeSet sLineDef, lLineOpenFrom, False, 3
    Debug.Assert sLineDef = "1-1"
    pModel.AppendToRangeSet sLineDef, lLineOpenFrom, True, 4
    Debug.Assert sLineDef = "1-1;4"
    pModel.AppendToRangeSet sLineDef, lLineOpenFrom, True, 5
    Debug.Assert sLineDef = "1-1;4"
    pModel.AppendToRangeSet sLineDef, lLineOpenFrom, False, 6
    Debug.Assert sLineDef = "1-1;4-5"
End Sub

Private Sub Test_StringBuilder()
    Dim sb As StringBuilder
    Set sb = New StringBuilder
    Debug.Assert Len(sb.ToString()) = 0
    sb.Append("This ").Append("is ").Append("a ").Append("test ").Append("of ").Append "the "
    sb.Append("StringBuilder's ").Append("incremental ").Append("expansion ").Append "ability."
    Debug.Assert sb.ToString() = "This is a test of the StringBuilder's incremental expansion ability."
    sb.Append vbNullString
    sb.Append ""
    Debug.Assert sb.ToString() = "This is a test of the StringBuilder's incremental expansion ability."
    sb.Append " "
    sb.Append "Now I'm adding a very long string to test that StringBuilder correctly handles strings that are " & _
            "more than double the length of the current buffer. Godspeed, StringBuilder! Lorem ipsum dolor sit " & _
            "amet, consectetur adipiscing elit. Suspendisse hendrerit lectus ligula, sodales rhoncus nunc " & _
            "porttitor vitae. Integer commodo vestibulum suscipit. Donec ultrices tellus ac tincidunt condimentum."
    Debug.Assert sb.ToString() = "This is a test of the StringBuilder's incremental expansion ability. " & _
            "Now I'm adding a very long string to test that StringBuilder correctly handles strings that are " & _
            "more than double the length of the current buffer. Godspeed, StringBuilder! Lorem ipsum dolor sit " & _
            "amet, consectetur adipiscing elit. Suspendisse hendrerit lectus ligula, sodales rhoncus nunc " & _
            "porttitor vitae. Integer commodo vestibulum suscipit. Donec ultrices tellus ac tincidunt condimentum."
    Set sb = New StringBuilder
    Debug.Assert sb.Append("This is a test of adding a big string to StringBuilder up-front. Will it choke? " & _
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse hendrerit lectus ligula, " & _
            "sodales rhoncus nunc porttitor vitae. Integer commodo vestibulum suscipit. Donec ultrices tellus " & _
            "ac tincidunt condimentum. Etiam volutpat ligula ipsum, a commodo neque tempor vitae. Vestibulum a " & _
            "cursus nisl. Interdum et malesuada fames ac ante ipsum primis in faucibus.").ToString() = _
            "This is a test of adding a big string to StringBuilder up-front. Will it choke? " & _
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse hendrerit lectus ligula, " & _
            "sodales rhoncus nunc porttitor vitae. Integer commodo vestibulum suscipit. Donec ultrices tellus " & _
            "ac tincidunt condimentum. Etiam volutpat ligula ipsum, a commodo neque tempor vitae. Vestibulum a " & _
            "cursus nisl. Interdum et malesuada fames ac ante ipsum primis in faucibus."
End Sub
