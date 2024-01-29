Sub Stock_Market()
    Dim ws As Worksheet
    Dim WorksheetName As String
    Dim Ticker_Name As String
    Dim Total_Stock_Volume As Variant
    Dim Yearly_Change As Variant
    Dim Percent_Change As Variant
    Dim summaryindex As Variant
    Dim MaxVal1 As Double
    Dim MinVal1 As Double
    Dim MaxVal2 As Double
    Dim MinVal2 As Double
    Dim Rng1 As Variant
    Dim Rng2 As Variant
    For Each ws In Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        Total_Stock_Volume = 0
        summaryindex = 0
        For I = 2 To LastRow
        
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                Ticker_Name = ws.Cells(I, 1).Value
                Yearly_Change = ws.Range("F" & 2 + summaryindex) - ws.Range("C" & 2 + summaryindex).Value
                Percent_Change = Round(100 * (Yearly_Change) / (ws.Cells(I, 3).Value), 2)
                 If Yearly_Change > 0 Then
                        ws.Range("J" & 2 + summaryindex).Interior.ColorIndex = 4
                    ElseIf Yearly_Change < 0 Then
                        ws.Range("J" & 2 + summaryindex).Interior.ColorIndex = 3
                        End If
                        
                'Output)
                ws.Range("I" & 2 + summaryindex).Value = Ticker_Name
                ws.Range("J" & 2 + summaryindex).Value = Yearly_Change
                ws.Range("K" & 2 + summaryindex).Value = Percent_Change & "%"
                ws.Range("L" & 2 + summaryindex).Value = Total_Stock_Volume
                
                Set Rng1 = ws.Range("K2:K" & LastRow)
                Set Rng2 = ws.Range("L2:L" & LastRow)
                
                'Locate the maximum and minimum values of Percent Change and paste the values in the appropriate cells
                MaxVal1 = WorksheetFunction.Max(Rng1)
                MinVal1 = WorksheetFunction.Min(Rng1)
                MaxVal2 = WorksheetFunction.Max(Rng2)
                
                'Paste the associated Ticker names with the highest and lowest percent changes and paste them into the appropriate cells
                If ws.Range("K" & 2 + summaryindex).Value = MaxVal1 Then
                    ws.Range("P" & 2).Value = Ticker_Name
                    ws.Range("Q" & 2).Value = MaxVal1
                ElseIf ws.Range("K" & 2 + summaryindex).Value = MinVal1 Then
                    ws.Range("P" & 3).Value = Ticker_Name
                    ws.Range("Q" & 3).Value = MinVal1
                     End If
                     
                If ws.Range("L" & 2 + summaryindex).Value = MaxVal2 Then
                    ws.Range("P" & 4).Value = Ticker_Name
                    ws.Range("Q" & 4).Value = Total_Stock_Volume
                    Else
                    End If
              
                'Reset Variables
                Total_Stock_Volume = 0
                summaryindex = summaryindex + 1
                
            Else
                Total_Stock_Volume = Total_Stock_Volume + ws.Range("G" & I).Value
                
                End If
                
        Next I
        Next ws
    
End Sub