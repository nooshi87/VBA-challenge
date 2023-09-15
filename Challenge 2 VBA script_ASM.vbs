Sub RunMultipleYeAR()


Dim ws As Worksheet

'setting row counter for FOR loops
Dim i As Double
Dim counter As Long
Dim LastRow As Long
Dim Ticker_Abr As String
'Setting variables for each column in summary table
Dim initialprice As Double
Dim closeprice As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock As Double


'Setting Summary table row counter
Dim Summary_Row As Integer

'Creating identifiers for each sheet
'Creating loop to pull last row from each sheet
    For Each ws In Worksheets
    Summary_Row = 2

       LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Total_Stock = 0
        counter = 0
 ws.Cells(1, 9).Value = "Ticker"
 ws.Cells(1, 10).Value = "Yearly Change"
 ws.Cells(1, 11).Value = "Percent Change"
 ws.Cells(1, 12).Value = "Total Stock Volume"
 
 
        For i = 2 To LastRow
    
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
             'pull ticker abbreviation into summary table
            Ticker_Abr = ws.Cells(i, 1).Value
                    ws.Cells(Summary_Row, 9).Value = Ticker_Abr
                 'Calculate yearly change by looking at first day open and last day close values
                    closeprice = ws.Cells(i, 6).Value
                    initialprice = ws.Cells(i - counter, 3).Value
                     Yearly_Change = closeprice - initialprice
                    ws.Cells(Summary_Row, 10).Value = Yearly_Change
                    If Yearly_Change > 0 Then
                    ws.Cells(Summary_Row, 10).Interior.ColorIndex = 4
                    ElseIf Yearly_Change < 0 Then
                    ws.Cells(Summary_Row, 10).Interior.ColorIndex = 3
                    Else
                     ws.Cells(Summary_Row, 10).Interior.ColorIndex = 0
                    End If
                    
                'Calculate percent change and format column K to reflect values
                    Percent_Change = Yearly_Change / initialprice
                       ws.Cells(Summary_Row, 11).Value = Percent_Change
                       ws.Range("K" & Summary_Row).NumberFormat = "0.00%"
                  'Calculate total stock volume for each ticker and insert in summary table
                    Total_Stock = Total_Stock + ws.Cells(i, 7).Value
                    ws.Cells(Summary_Row, 12).Value = Total_Stock
                    
                Total_Stock = 0
                Summary_Row = Summary_Row + 1
                counter = 0
            
      Else
        counter = counter + 1
        Total_Stock = Total_Stock + ws.Cells(i, 7).Value

      End If
      Next i
      
'Creating a second summary table
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest Total volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
'Setting variables for second summary table
    Dim Summary2row As Integer
    Dim Match1 As Double
    Dim match2 As Double
    Dim match3 As Double
    Summary2row = 2
    ws.Cells(Summary2row, 17).Value = WorksheetFunction.Max(Range("k2:k" & LastRow))
'find position of max% diff, min% diff, and greatest volume values. Insert respective tickers and identified values to summary table 2
    Match1 = WorksheetFunction.Match(WorksheetFunction.Max(Range("k2:k" & LastRow)), Range("k2:k" & LastRow), 0)
    ws.Cells(Summary2row, 16).Value = ws.Cells(Match1 + 1, 9).Value
    ws.Cells(Summary2row, 17).NumberFormat = "0.00%"
    
    
    Summary2row = Summary2row + 1
    ws.Cells(Summary2row, 17).Value = WorksheetFunction.Min(Range("k2:k" & LastRow))
    match2 = WorksheetFunction.Match(WorksheetFunction.Min(Range("k2:k" & LastRow)), Range("k2:k" & LastRow), 0)
    ws.Cells(Summary2row, 16).Value = ws.Cells(match2 + 1, 9).Value
    ws.Cells(Summary2row, 17).NumberFormat = "0.00%"

    Summary2row = Summary2row + 1
    ws.Cells(Summary2row, 17).Value = WorksheetFunction.Max(Range("l2:l" & LastRow))
    match3 = WorksheetFunction.Match(WorksheetFunction.Max(Range("l2:l" & LastRow)), Range("l2:l" & LastRow), 0)
    ws.Cells(Summary2row, 16).Value = ws.Cells(match3 + 1, 9).Value

  
    'Autofit column width so contents per cell are visible
    ws.Columns("A:P").AutoFit

    Next ws
        
End Sub




