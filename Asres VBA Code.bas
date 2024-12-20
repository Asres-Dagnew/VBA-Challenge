Attribute VB_Name = "Module1"
Sub QuarterlyAnalysis()
    ' Declare variables for tracking and summary
    Dim ws As Worksheet
    Dim ticker As String
    Dim openPrice As Double, closePrice As Double, percentageChange As Double, totalStockVol As Double
    Dim quarterlyChange As Double, greatestTotalVol As Double
    Dim lastRow As Long
    Dim previousStockPrice As Long, summaryRow As Integer
    Dim greatestIncrease As Double, greatestDecrease As Double

    For Each ws In ThisWorkbook.Worksheets
        ws.Activate

        ' Initialize variables
        totalStockVol = 0
        greatestTotalVol = 0
        previousStockPrice = 2
        summaryRow = 2
        greatestIncrease = 0
        greatestDecrease = 0

        ' Set column headers
        ws.Range("P1:Q1").Value = Array("Ticker", "Value")
        ws.Range("O2:O4").Value = Application.Transpose(Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"))
        ws.Range("I1:L1").Value = Array("Ticker", "Quarterly Change", "Percent Change", "Total Stock Volume")

        ' Get the last row of data
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Loop through rows
        For i = 2 To lastRow
            totalStockVol = totalStockVol + ws.Cells(i, 7).Value

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(previousStockPrice, 3).Value
                closePrice = ws.Cells(i, 6).Value
                quarterlyChange = closePrice - openPrice

                If openPrice = 0 Then
                    percentageChange = 0
                Else
                    percentageChange = quarterlyChange / openPrice
                End If

                ' Record values in the summary table
                With ws
                    .Cells(summaryRow, 9).Value = ticker
                    .Cells(summaryRow, 10).Value = quarterlyChange
                    .Cells(summaryRow, 11).Value = percentageChange
                    .Cells(summaryRow, 12).Value = totalStockVol
                    .Cells(summaryRow, 11).NumberFormat = "0.00%"
                    
            'Use conditional formatting that will highlight positive change in green and negative change in red.
                    If quarterlyChange > 0 Then
                        .Cells(summaryRow, 10).Interior.ColorIndex = 4 ' Green for positive change
                    ElseIf quarterlyChange < 0 Then
                        .Cells(summaryRow, 10).Interior.ColorIndex = 3 ' Red for negative change
                    Else
                        .Cells(summaryRow, 10).Interior.ColorIndex = 0 ' No color for no change
                    End If
                End With

                ' Reset for next stock
                totalStockVol = 0
                summaryRow = summaryRow + 1
                previousStockPrice = i + 1
            End If
        Next i

        ' Find greatest values
        For i = 2 To ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
            If ws.Cells(i, 12).Value > greatestTotalVol Then
                greatestTotalVol = ws.Cells(i, 12).Value
                ws.Cells(4, 17).Value = greatestTotalVol
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            End If
            If ws.Cells(i, 11).Value > greatestIncrease Then
                greatestIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 17).Value = greatestIncrease
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            End If
            If ws.Cells(i, 11).Value < greatestDecrease Then
                greatestDecrease = ws.Cells(i, 11).Value
                ws.Cells(3, 17).Value = greatestDecrease
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            End If
        Next i

        ' Format greatest percentage values as percentage
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Columns("O:O").EntireColumn.AutoFit
        ws.Columns("L:L").EntireColumn.AutoFit
        ws.Columns("J:J").EntireColumn.AutoFit
         ws.Columns("J:J").NumberFormat = "0.00"
        ws.Columns("K:K").EntireColumn.AutoFit
    Next ws
End Sub


