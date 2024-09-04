Attribute VB_Name = "Module1"
Sub StockInfo()
    ' Declare variables to be used in the subroutine
    Dim ws As Worksheet

    Dim ticker As String
    Dim prevTicker As String
    Dim quarterOpenValue As Double
    Dim dailyClosingValue As Double

    Dim i As Long
    Dim k As Long

    Dim startingValue As Double
    Dim endingValue As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim stockVolume As Double
    Dim stockIndex As Long
    Dim totalVolume As Double
    Dim maxInc As Double
    Dim maxIncTicker As String
    Dim maxDec As Double
    Dim maxDecTicker As String
    Dim maxVolume As Double
    Dim maxVolumeTicker As String
    

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Step 1: Create headings in the summary columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Volume"
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ' Step 2: Initialize the first ticker and starting value
        ticker = ws.Cells(2, 1).Value
        ws.Cells(2, 9).Value = ticker
        startingValue = ws.Cells(2, 3).Value
        
        ' Initialize loop variables
        i = 3
        stockIndex = 2
        
        ' Step 3: Loop through all the rows in the worksheet
        While ws.Cells(i, 1).Value <> ""
            ' Set the previous ticker to the current ticker
            prevTicker = ticker
            ' Get the current ticker from the cell
            ticker = ws.Cells(i, 1).Value
            
            ' Check if the ticker has changed (new stock)
            If prevTicker <> ticker Then
                ' Step 4: Calculate the ending value and changes
                endingValue = ws.Cells(i - 1, 6).Value
                quarterlyChange = startingValue - endingValue
                percentChange = quarterlyChange / endingValue
                
                ' Compare volume and change values
                If percentChange > maxInc Then
                    maxInc = percentChange
                    maxIncTicker = prevTicker
                ElseIf percentChange < maxDec Then
                    maxDec = percentChange
                    maxDecTicker = prevTicker
                End If
                
                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    maxVolumeTicker = prevTicker
                End If
                ' Step 5: Print the summary for the previous ticker
                ws.Cells(stockIndex, 9).Value = prevTicker
                ws.Cells(stockIndex, 10).Value = quarterlyChange
                ws.Cells(stockIndex, 11).Value = percentChange
                ws.Cells(stockIndex, 12).Value = totalVolume
                
                ' Reset the total volume and increment the stock index
                totalVolume = 0
                stockIndex = stockIndex + 1
                
                ' Reinitialize starting value for the new ticker
                startingValue = ws.Cells(i, 3).Value
            Else
                ' Step 6: Accumulate the stock volume for the current ticker
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
            
            ' Move to the next row
            i = i + 1
        Wend
        
        ' Step 7: Report Changes
        ws.Cells(2, 16).Value = maxIncTicker
        ws.Cells(2, 17).Value = maxInc
        
        ws.Cells(3, 16).Value = maxDecTicker
        ws.Cells(3, 17).Value = maxDec
        
        ws.Cells(4, 16).Value = maxVolumeTicker
        ws.Cells(4, 17).Value = maxVolume
        
        ' Step 8: Conditional Formatting
        For k = 2 To stockIndex
            ' Set losses to Red
            If ws.Cells(k, 11).Value < 0 Then
                ws.Cells(k, 11).Interior.Color = RGB(255, 0, 0)
            ' Set gains to Green
            ElseIf ws.Cells(k, 11).Value > 0 Then
                ws.Cells(k, 11).Interior.Color = RGB(0, 255, 0)
            End If
        Next k
        
      
        ' Step 9: Number Formatting
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("P2:Q3").NumberFormat = "0.00%"
            
    Next ws
End Sub



