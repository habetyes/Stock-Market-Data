Sub Stock()

Dim Ticker As String
Dim Summaryrow As Integer
Dim TotalVolume As Double

'Loop through each worksheet
For Each ws In Worksheets

'Set Column Titles and initial summary row for each ws
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Total Stock Volume"
Summaryrow = 2
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Loop Through the all data rows
    For i = 2 To lastrow
    
        'For the last occurance of a ticker symbol
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'Assign the ticker name and total volume to those variables
        Ticker = ws.Cells(i, 1).Value
        TotalVolume = TotalVolume + ws.Cells(i, 7)
        
        'Place the values for Name and Volume in the appropriate cells
        ws.Range("I" & Summaryrow).Value = Ticker
        ws.Range("J" & Summaryrow).Value = TotalVolume
        
        'Place the Name and Total on the next row
        Summaryrow = Summaryrow + 1
        
         TotalVolume = 0
        
        
        Else
            'Keep a running tab of total volume for each ticker symbol
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
        End If
        
    Next i
    
Next ws

End Sub
