Sub VBA_challenge()

'Set the initial variables'

Dim ws As Worksheet

Dim Ticker As String
Dim Total_Stock_Volume As Double
Dim Summary_Table_Row As Integer
Dim First_Opening_Price As Double
Dim Last_Closing_Price As Double
Dim maxIncrease As Double
Dim minIncrease As Double
Dim maxTotalVolume As Double

For Each ws In Worksheets

    Total_Stock_Volume = 0
    Summary_Table_Row = 2
    
'Inserting Data Via Cells
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        ws.Cells(2, 16).Value = "Greatest % increase"
        ws.Cells(3, 16).Value = "Greatest % decrease"
        ws.Cells(4, 16).Value = "Greatest total volume"

'Loop through all tickers
        For i = 2 To 753001

'Check if the next ticker is different than the current one
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
'Inform the ticker symbol
                Ticker = ws.Cells(i, 1).Value
    
'Add to the total stock volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
'Print the Ticker Name in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                Last_Closing_Price = ws.Cells(i, 6).Value
            
'Print the Yearly change in the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = Last_Closing_Price - First_Opening_Price
            
'Print the Percentage change in the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = (Last_Closing_Price - First_Opening_Price) / First_Opening_Price
            
'Print the Volume to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        
                Summary_Table_Row = Summary_Table_Row + 1
        
                Total_Stock_Volume = 0
            Else
            
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                First_Opening_Price = ws.Cells(i, 3).Value
            
            End If
            
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            
    
        End If
        
    Next i

'Find the greatest values
    maxIncrease = Application.WorksheetFunction.Max(ws.Range("K2:K3001"))
    ws.Cells(2, 18) = maxIncrease
    ws.Cells(2, 18).NumberFormat = "0.00%"
    minIncrease = Application.WorksheetFunction.Min(ws.Range("K2:K3001"))
    ws.Cells(3, 18) = minIncrease
    ws.Cells(3, 18).NumberFormat = "0.00%"
    maxTotalVolume = Application.WorksheetFunction.Max(ws.Range("L2:L3001"))
    ws.Cells(4, 18) = maxTotalVolume
    
'Find the ticker associated with greatest values
    maxRow = Application.WorksheetFunction.Match(maxIncrease, ws.Range("K2:K3001"), 0)
    ws.Cells(2, 17) = ws.Cells(maxRow + 1, 9).Value

    minRow = Application.WorksheetFunction.Match(minIncrease, ws.Range("K2:K3001"), 0)
    ws.Cells(3, 17) = ws.Cells(minRow + 1, 9).Value

    maxTotalVolumeRow = Application.WorksheetFunction.Match(maxTotalVolume, ws.Range("L2:L3001"), 0)
    ws.Cells(4, 17) = ws.Cells(maxTotalVolumeRow + 1, 9).Value

' Formatting
    For i = 2 To 3001
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
        
        End If
    
    
    ws.Cells(i, 11).NumberFormat = "0.00%"
        
Next i
    
Next

End Sub

