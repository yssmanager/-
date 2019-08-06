Attribute VB_Name = "Module1"
Sub selectLatest()

    Dim MaxRow As Long
    Dim MaxCol As Long
    Dim ws As Worksheet
    
    Set ws = Sheets("Sheet2")
        
    Call lastRC(ws, MaxRow, MaxCol)
    
    ws.Cells(MaxRow, 1).Activate

End Sub

Function lastRC(ws As Worksheet, Row As Long, Col As Long)

    With ws.UsedRange
        Row = .Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
        Col = .Find("*", , xlFormulas, , xlByColumns, xlPrevious).Column
    End With

End Function

