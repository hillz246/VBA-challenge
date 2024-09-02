Sub stock_data()

    Dim ws As Worksheet
    Dim last_Row As Long
    Dim ticker As String
    Dim open_price As Double
    Dim close_price As Double
    Dim quarterly_change As Double
    Dim Percentage_Change As Double
    Dim total_volume As Double
    Dim row As Long
    Dim column As Integer
    Dim i As Long

    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        last_Row = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        total_volume = 0
        row = 2
        column = 1
        
        If IsNumeric(ws.Cells(2, column + 2).Value) Then
        open_price = ws.Cells(2, column + 2).Value
        Else: open_price = 0
        End If

        For i = 2 To last_Row
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
                ticker = ws.Cells(i, column).Value
                ws.Cells(row, column + 8).Value = ticker
                close_price = ws.Cells(i, column + 5).Value
                quarterly_change = close_price - open_price
                ws.Cells(row, column + 9).Value = quarterly_change

                If open_price <> 0 Then
                    Percentage_Change = quarterly_change / open_price
                Else
                    Percentage_Change = 0
                End If

                ws.Cells(row, column + 10).Value = Percentage_Change
                ws.Cells(row, column + 10).NumberFormat = "0.00%"
                total_volume = total_volume + ws.Cells(i, column + 6).Value
                ws.Cells(row, column + 11).Value = total_volume
                row = row + 1
                open_price = ws.Cells(i + 1, column + 2).Value
                total_volume = 0
            Else
                total_volume = total_volume + ws.Cells(i, column + 6).Value
            End If
        Next i

        quarterly_change_last_row = ws.Cells(ws.Rows.Count, 9).End(xlUp).row
        For j = 2 To quarterly_change_last_row
            If ws.Cells(j, 10).Value > 0 Or ws.Cells(j, 10).Value = 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 10
            ElseIf ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j

        ws.Cells(1, 20).Value = "Ticker"
        ws.Cells(1, 21).Value = "Value"
        ws.Cells(2, 19).Value = "Greatest % increase"
        ws.Cells(3, 19).Value = "Greatest % decrease"
        ws.Cells(4, 19).Value = "Greatest total volume"
   
   
   
        Dim Greatest_percentage_increase As Double
        Dim Greatest_percentage_decrease As Double
        Dim Greatest_total_volume As Double
        

' Find the max and min percentage changes
        Greatest_percentage_increase = Application.WorksheetFunction.Max(ws.Range("K2:K" & quarterly_change_last_row))
        Greatest_percentage_decrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & quarterly_change_last_row))
        Greatest_total_volume = Application.WorksheetFunction.Max(ws.Range("L2:L" & quarterly_change_last_row))

        For k = 2 To quarterly_change_last_row
            If ws.Cells(k, 11).Value = Greatest_percentage_increase Then
               ws.Cells(2, 20).Value = ws.Cells(k, 9).Value
               ws.Cells(2, 21).Value = ws.Cells(k, 11).Value
               ws.Cells(2, 21).NumberFormat = "0.00%"
            End If
    
            If ws.Cells(k, 11).Value = Greatest_percentage_decrease Then
               ws.Cells(3, 20).Value = ws.Cells(k, 9).Value
               ws.Cells(3, 21).Value = ws.Cells(k, 11).Value
               ws.Cells(3, 21).NumberFormat = "0.00%"
            End If
    
            If ws.Cells(k, 12).Value = Greatest_total_volume Then
               ws.Cells(4, 20).Value = ws.Cells(k, 9).Value
               ws.Cells(4, 21).Value = ws.Cells(k, 12).Value
             End If
             
        Next k
    Next ws
End Sub



