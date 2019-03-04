Sub Stock()

Dim Ticker As String
Dim Summaryrow As Integer
Dim TotalVolume As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim OpenPrice As Double
Dim ClosingPrice As Double

'Loop through each worksheet
For Each ws In Worksheets

'Set Column Titles, initial summary row and lastrow for each ws
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
Summaryrow = 2
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Store Percent Change Formula As Variable

    'Loop Through the all data rows
    For i = 2 To lastrow
    
        'Store Open Price as Variable
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        OpenPrice = ws.Cells(i, 3).Value
        
        End If
        
            'For the last occurance of a ticker symbol
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Store Closing Price as Variable
            ClosingPrice = ws.Cells(i, 6).Value
            
            'Assign the ticker name and total volume to those variables
            Ticker = ws.Cells(i, 1).Value
            TotalVolume = TotalVolume + ws.Cells(i, 7)
            
            'Place the values for Name and Volume in the appropriate cells
            ws.Range("I" & Summaryrow).Value = Ticker
            ws.Range("L" & Summaryrow).Value = TotalVolume
            
            'Store Yearly Change as Variables
            YearlyChange = ClosingPrice - OpenPrice
            
              'Store Percent Change as Variable and avoid divide by 0 error
                    If OpenPrice = 0 Then
                    PercentChange = 0
                    Else: PercentChange = (ClosingPrice - OpenPrice) / OpenPrice
                    End If
                    
                'Conditionally Format Cell
                If YearlyChange > 0 Then
                ws.Range("J" & Summaryrow).Interior.ColorIndex = 4
                
                ElseIf YearlyChange < 0 Then
                ws.Range("J" & Summaryrow).Interior.ColorIndex = 3
                
                Else: ws.Range("J" & Summaryrow).Interior.ColorIndex = 2
                             
                End If
                
            
            'Place Yearly and Percent changes in appropriate cells
            ws.Range("J" & Summaryrow).Value = YearlyChange
            ws.Range("K" & Summaryrow).Value = PercentChange
            
            
            'Place the Name and Total on the next row
            Summaryrow = Summaryrow + 1
            
            'Reset Total Volume Running Total
             TotalVolume = 0
            
            
            Else
                'Keep a running tab of total volume for each ticker symbol
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
            End If
        
        
    Next i
    

Sub Stock()

Dim Ticker As String
Dim Summaryrow As Integer
Dim TotalVolume As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim OpenPrice As Double
Dim ClosingPrice As Double

'Loop through each worksheet
For Each ws In Worksheets

'Set Column Titles, initial summary row and lastrow for each ws
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
Summaryrow = 2
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Store Percent Change Formula As Variable

    'Loop Through the all data rows
    For i = 2 To lastrow
    
        'Store Open Price as Variable
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        OpenPrice = ws.Cells(i, 3).Value
        
        End If
        
            'For the last occurance of a ticker symbol
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Store Closing Price as Variable
            ClosingPrice = ws.Cells(i, 6).Value
            
            'Assign the ticker name and total volume to those variables
            Ticker = ws.Cells(i, 1).Value
            TotalVolume = TotalVolume + ws.Cells(i, 7)
            
            'Place the values for Name and Volume in the appropriate cells
            ws.Range("I" & Summaryrow).Value = Ticker
            ws.Range("L" & Summaryrow).Value = TotalVolume
            
            'Store Yearly Change as Variables
            YearlyChange = ClosingPrice - OpenPrice
            
              'Store Percent Change as Variable and avoid divide by 0 error
                    If OpenPrice = 0 Then
                    PercentChange = 0
                    Else: PercentChange = (ClosingPrice - OpenPrice) / OpenPrice
                    End If
                    
                'Conditionally Format Cell
                If YearlyChange > 0 Then
                ws.Range("J" & Summaryrow).Interior.ColorIndex = 4
                
                ElseIf YearlyChange < 0 Then
                ws.Range("J" & Summaryrow).Interior.ColorIndex = 3
                
                Else: ws.Range("J" & Summaryrow).Interior.ColorIndex = 2
                             
                End If
                
            
            'Place Yearly and Percent changes in appropriate cells
            ws.Range("J" & Summaryrow).Value = YearlyChange
            ws.Range("K" & Summaryrow).Value = PercentChange
            
            
            'Place the Name and Total on the next row
            Summaryrow = Summaryrow + 1
            
            'Reset Total Volume Running Total
             TotalVolume = 0
            
            
            Else
                'Keep a running tab of total volume for each ticker symbol
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
            End If
        
        
    Next i

Range("K:K").NumberFormat = "0.00%"   
Next ws

End Sub

Next ws

End Sub

