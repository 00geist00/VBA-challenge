Sub Stock_Summary_AllSheets()
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate

        ' Add Heading for summary
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        'Create and define variables
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Ticker_Name As String
        Dim Percent_Change As Double
        Dim StockVolumn As Double
        StockVolumn = 0
        Dim Row As Integer
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
        
        'Set Initial Open Price
        Open_Price = Cells(Row, Column + 2).Value
         ' Loop through all ticker symbol
        
        ' Determine the Last Row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow
         ' Check if we are still within the same ticker symbol (must be sorted by ticker symbol and open date)
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ' Set Ticker name in Summary
                Ticker_Name = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = Ticker_Name
                ' Set Close Price for Summary
                Close_Price = Cells(i, Column + 5).Value
                ' Add Yearly Change
                Yearly_Change = Close_Price - Open_Price
                Cells(Row, Column + 9).Value = Yearly_Change
                ' Add Percent Change
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(Row, Column + 10).Value = Percent_Change
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                ' Add Total StockVolumn
                StockVolumn = StockVolumn + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = StockVolumn
                Row = Row + 1
                ' reset the Open Price
                Open_Price = Cells(i + 1, Column + 2)
                ' reset the StockVolumn Total
                StockVolumn = 0
                
            'if cells are the same ticker add to StockVolumn
            Else
                StockVolumn = StockVolumn + Cells(i, Column + 6).Value
            End If
        Next i

        ' Conditional formatting for Yearly Change column
        For j = 2 To LastRow
            If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
        
      Next WS
End Sub