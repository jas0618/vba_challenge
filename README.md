The Multiple_year_stock_data excel shows the quaterly change and percent change(increase or decrease) of each stock. And which stock had the greatest percentage increase or decrease and the greatest volume. 

The first step is to define the variables 
Second Step to indicate which columns/rows have the opening price and closing price
Third: calculate the quaterly change and percent change(between the opening price and closing price)
Fourth: Creating a loop(conditional) in the case the percent change is less or more zero
*following the same pathway for the alphabetical_testing 
Please find below the comments with the script:


Sub AnalyzeStocks()
The following are the variables:
    Dim ws As Worksheet
    Dim i As Long, lastrow As Long
    Dim summary_table_row As Long
    Dim opening_price As Double, closing_price As Double
    Dim total_volume As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestvolume As Double
    Dim jIncrease As String
    Dim jDecrease As String
    Dim jvolume As String
    Dim quarterly_change As Double, percent_change As Double

    greatestIncrease = -999999
    greatestDecrease = 999999
    greatestvolume = 0

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row in Column A of the current worksheet
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        summary_table_row = 4 ' Starting row for summary table (adjust as needed)

        ' Loop through all rows in the current worksheet
        For i = 2 To lastrow
            ' Check if the ticker changes (i.e., new ticker or first row)
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                opening_price = ws.Cells(i, 3).Value ' Set opening price
                total_volume = 0 ' Reset total volume
            End If

            ' Accumulate total volume
            total_volume = total_volume + ws.Cells(i, 7).Value

            ' Check if the ticker changes on the next row (i.e., last row for current ticker)
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Get the closing price
                closing_price = ws.Cells(i, 6).Value

                ' Calculate quarterly change and percent change
                quarterly_change = closing_price - opening_price
                If opening_price <> 0 Then
                    percent_change = ((closing_price - opening_price) / opening_price) * 100
                Else
                    percent_change = 0
                End If

                ' Track greatest decrease
                If percent_change < greatestDecrease Then
                    greatestDecrease = percent_change
                    jDecrease = ws.Cells(i, 1).Value
                End If

                ' Track greatest increase
                If percent_change > greatestIncrease Then
                    greatestIncrease = percent_change
                    jIncrease = ws.Cells(i, 1).Value
                End If

                ' Track greatest volume
                If total_volume > greatestvolume Then
                    greatestvolume = total_volume
                    jvolume = ws.Cells(i, 1).Value
                End If

                ' Output the values to the summary table
                ws.Cells(summary_table_row, 10).Value = ws.Cells(i, 1).Value ' Ticker in Column J
                ws.Cells(summary_table_row, 11).Value = quarterly_change ' Quarterly change in Column K
                ws.Cells(summary_table_row, 12).Value = percent_change ' Percent change in Column L
                ws.Cells(summary_table_row, 12).Value = Format(percent_change, "0.00") & "%"
                ws.Cells(summary_table_row, 13).Value = total_volume ' Total volume in Column M

                ' Move to the next summary table row
                summary_table_row = summary_table_row + 1
            End If
        Next i

        ' Color the cells based on the value in Column K (Quarterly Change)
        For i = 4 To ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
            Select Case ws.Cells(i, "K").Value
                Case Is > 0
                    ws.Cells(i, "K").Interior.ColorIndex = 4 ' Green for positive change
                Case Is < 0
                    ws.Cells(i, "K").Interior.ColorIndex = 3 ' Red for negative change
                Case Else
                    ws.Cells(i, "K").Interior.ColorIndex = 0 ' No fill for zero change
            End Select
        Next i

        ' Output summary table for greatest increase, decrease, and volume starting from P2
        ws.Cells(2, 16).Value = "Summary" ' Column P2
        ws.Cells(3, 16).Value = "Greatest % Increase"
        ws.Cells(3, 17).Value = jIncrease ' Column Q3
        ws.Cells(3, 18).Value = Format(greatestIncrease, "0.00") & "%" ' Column R3

        ws.Cells(4, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 17).Value = jDecrease ' Column Q4
        ws.Cells(4, 18).Value = Format(greatestDecrease, "0.00") & "%" ' Column R4

        ws.Cells(5, 16).Value = "Greatest Volume"
        ws.Cells(5, 17).Value = jvolume ' Column Q5
        ws.Cells(5, 18).Value = greatestvolume ' Column R5

    Next ws

End Sub




