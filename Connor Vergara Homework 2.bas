Attribute VB_Name = "Module1"
Sub StockBreaker()

Dim ws As Worksheet

    For Each ws In Sheets
    
    Dim Stock_Ticker As String
    
    Dim Stock_Total As Double
    Stock_Total = 0
    
    Dim Stock_Open As Double
    Dim Stock_Close As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    
    Dim LastRow As Double
    Dim RowHolder As Double
    RowHolder = 1
    
    Dim Stock_Row As Double
    Stock_Row = 2
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Columns("J").ColumnWidth = 13
    ws.Columns("K").ColumnWidth = 13
    ws.Columns("L").ColumnWidth = 16
    ws.Columns("O").ColumnWidth = 20
    ws.Columns("Q").ColumnWidth = 15
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            Stock_Ticker = ws.Cells(i, 1).Value
            
            Stock_Open = ws.Cells(RowHolder + 1, 3).Value
            
            RowHolder = i
            
            Stock_Close = ws.Cells(i, 6).Value
            
            Stock_Total = Stock_Total + ws.Cells(i, 7).Value
            
            Yearly_Change = Stock_Close - Stock_Open
            
            If Yearly_Change = 0 Or Stock_Open = 0 Then
        
                Percent_Change = 0
                
                Stock_Open = 1
                
                Else
            
                Percent_Change = (Yearly_Change / Stock_Open)
                
            End If
        
            ws.Range("I" & Stock_Row).Value = Stock_Ticker
            
                If Yearly_Change > 0 Then
                    
                    ws.Range("J" & Stock_Row).Value = Yearly_Change
                    ws.Range("J" & Stock_Row).Interior.ColorIndex = 4
                    
                ElseIf Yearly_Change < 0 Then
                
                    ws.Range("J" & Stock_Row).Value = Yearly_Change
                    ws.Range("J" & Stock_Row).Interior.ColorIndex = 3
                    
                Else
                
                    ws.Range("J" & Stock_Row).Value = Yearly_Change
                    
                End If
               
            ws.Range("K" & Stock_Row).Value = Percent_Change
            ws.Range("K" & Stock_Row).NumberFormat = "0.00%"
            
            ws.Range("L" & Stock_Row).Value = Stock_Total
            
            Stock_Row = Stock_Row + 1
            
            Stock_Total = 0
            
        Else
            
            Stock_Total = Stock_Total + ws.Cells(i, 7).Value
            
        End If
        
    Next i
    
    LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    Ticker2 = ws.Cells(2, 9).Value
    Ticker3 = ws.Cells(2, 9).Value
    Ticker4 = ws.Cells(2, 9).Value
    GreatestIncrease = ws.Cells(2, 11).Value
    GreatestDecrease = ws.Cells(2, 11).Value
    GreatestVolume = ws.Cells(2, 12).Value
    
    For i = 2 To LastRow2
    
        If ws.Cells(i + 1, 11).Value > GreatestIncrease Then
        
        GreatestIncrease = ws.Cells(i + 1, 11).Value
        Ticker2 = ws.Cells(i + 1, 9).Value
        
        End If
        
    Next i
    
    For i = 2 To LastRow2
    
        If ws.Cells(i + 1, 11).Value < GreatestDecrease Then
        
        GreatestDecrease = ws.Cells(i + 1, 11).Value
        Ticker3 = ws.Cells(i + 1, 9).Value
        
        End If
        
    Next i
    
    For i = 2 To LastRow2
    
        If ws.Cells(i + 1, 12).Value > GreatestVolume Then
        
        GreatestVolume = ws.Cells(i + 1, 12).Value
        Ticker4 = ws.Cells(i + 1, 9).Value
        
        End If
        
    Next i
    
    ws.Cells(2, 16).Value = Ticker2
    ws.Cells(2, 17).Value = GreatestIncrease
    ws.Cells(3, 16).Value = Ticker3
    ws.Cells(3, 17).Value = GreatestDecrease
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    ws.Cells(4, 16).Value = Ticker4
    ws.Cells(4, 17).Value = GreatestVolume

Next ws

End Sub
