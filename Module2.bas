Attribute VB_Name = "Module2"
Sub Summ_Table_Conditional()
    'run after Module 1

Dim ws As Worksheet

For Each ws In Worksheets

    ws.Columns("I").ColumnWidth = 12
    ws.Columns("J").ColumnWidth = 14
    ws.Columns("K").ColumnWidth = 14
    ws.Columns("L").ColumnWidth = 16
    LastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
      
      For j = 2 To LastRow
      
        If ws.Cells(j, 10).Value > 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 4
            Else
            ws.Cells(j, 10).Interior.ColorIndex = 3
        End If
        
        Next j
        Next ws
        
End Sub
