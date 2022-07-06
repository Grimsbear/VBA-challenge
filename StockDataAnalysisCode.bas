Attribute VB_Name = "Module1"
Sub StockDataAnalysis()
    
    'Big Thanks to Sanoo Singh & Scott Neubauer for helping me get through
    'Counter issue I had with these codes
    
    Dim Ticker As String
    
    Dim SummaryTable As Integer
    SummaryTable = 2
    
    Dim Volume As Variant
    Volume = 0
    
    Dim Counter As Integer
    Counter = 0
    
    For Each ws In Worksheets

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Counter = 0
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Annual Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Volume"
        
        For i = 2 To LastRow
           
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                'Print The Individual Ticker
                Ticker = Cells(i, 1).Value
                Cells(SummaryTable, 9).Value = Ticker
                
                'Print The Volume of Each Stock
                Volume = Volume + Cells(i, 7).Value
                Cells(SummaryTable, 12).Value = Volume
                
                'Find Value Changes
                StockOpen = Cells(i - (Counter - 1), 3).Value
                StockClose = Cells(i, 6).Value
                Cells(SummaryTable, 10).Value = (StockClose - StockOpen)
                Cells(SummaryTable, 11).Value = ((StockClose - StockOpen) / StockOpen)
                'Cells(SummaryTable, 13).Value = Counter
                 
                'Formatting Cells
                Cells(SummaryTable, 11).NumberFormat = "0.00%"
                
                    If Cells(SummaryTable, 10).Value >= 0 Then
                        
                        Cells(SummaryTable, 10).Interior.ColorIndex = 4
                    
                    Else
                    
                        Cells(SummaryTable, 10).Interior.ColorIndex = 3
                        
                    End If

                
                'Reset and Advance Variables
                Counter = 0
                StockOpen = 0
                StockClose = 0
                Volume = 0
                SummaryTable = SummaryTable + 1
            Else
                Counter = Counter + 1
                Volume = Volume + Cells(i, 7).Value
            End If
            
        Next i

        
    Next ws

End Sub
