Attribute VB_Name = "Module1"
Sub Stock_Market_Challenge()
    'run before Module 2
    
    'defining variables
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Volume As LongLong
    Dim Summary_Table_Row As Integer
    Dim Open_Value As Double
    Dim Close_Value As Double
    Dim Volume As LongLong
    Dim LastRow As Double
    Dim ws As Worksheet

For Each ws In Worksheets
        
    'creating headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 9).Font.Bold = True
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 10).Font.Bold = True
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 11).Font.Bold = True
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 12).Font.Bold = True
    
    ws.Columns("J").NumberFormat = "0.00"
    ws.Columns("K").NumberFormat = "0.00%"
    ws.Columns("L").NumberFormat = "0,000"
    
    'defining variable for last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'defining variable for the first row in the Summary Table
    Summary_Table_Row = 2
    'setting initial open value
    Open_Value = ws.Cells(2, 3).Value
    
    
    'loop through Ticker symbols
    For i = 2 To LastRow
        
        'check to see if we are on the same Ticker
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'add Ticker symbol to Summary Table
        Ticker = ws.Cells(i, 1).Value
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        
        'add total volume to Summary Table for each Ticker
        Volume = ws.Cells(i, 7).Value
        Total_Volume = Total_Volume + Volume
        ws.Range("L" & Summary_Table_Row).Value = Total_Volume
        Total_Volume = 0
        
        'add yearly change to Summary Table for each Ticker
        Close_Value = ws.Cells(i, 6).Value
        Yearly_Change = Close_Value - Open_Value
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
        
        'add percent change to Summary Table for each Ticker
        If Open_Value = 0 Then
            Percent_Change = 0
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            Else
            Percent_Change = (Yearly_Change / Open_Value)
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
         End If
        
        'move to next row to select open value for next Ticker
        Open_Value = ws.Cells(i + 1, 3).Value
    
    
        'move to next row in Summary Table
        Summary_Table_Row = Summary_Table_Row + 1
        
        
        
        Else
        
        'adding daily volumes for same Ticker
        Volume = ws.Cells(i, 7).Value
        Total_Volume = Total_Volume + Volume
    
      
        End If
        
        
        Next i
                                  
     
        
        
       
    Next ws
        
End Sub


    
     
   
