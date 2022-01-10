Attribute VB_Name = "Module1"
Sub StockAnalysis()
    
    
    Dim ws As Worksheet

    'Start loop
    For Each ws In Worksheets

        'Create column labels for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'Set variables
        Dim ticker_symbol As String
        Dim total_vol As Double
        total_vol = 0
        Dim rowcount As Long
        rowcount = 2
        Dim year_open As Double
        year_open = 0
        Dim year_close As Double
        year_close = 0
        Dim year_change As Double
        year_change = 0
        Dim percent_change As Double
        percent_change = 0
        Dim lastrow As Long
        
        'figure out final row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Loop to search ticker symbols
        For i = 2 To lastrow
            
            'Conditional to grab year open price
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

                year_open = ws.Cells(i, 3).Value

            End If

            'Total up the volume for each row to determine the total stock volume for the year
            total_vol = total_vol + ws.Cells(i, 7)

            'Conditional to determine if the ticker symbol is changing
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

                'Getting all the results
                ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value
                total_vol = ws.Cells(rowcount, 12).Value
                year_close = ws.Cells(i, 6).Value
                year_change = year_close - year_open
                ws.Cells(rowcount, 10).Value = year_change

                'Conditional to format coloring
                If year_change >= 0 Then
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 3
                End If

                'Calculate the percent change for the year and move it to the summary table format as a percentage
                If year_open = 0 And year_close = 0 Then
                    percent_change = 0
                    ws.Cells(rowcount, 11).Value = percent_change
                
                ElseIf year_open = 0 Then
                    'If a stock starts at zero and increases, it grows by infinite percent.
                    Dim percent_change_A As String
                    percent_change_A = "New Stock"
                    ws.Cells(rowcount, 11).Value = percent_change
                Else
                    percent_change = year_change / year_open
                    ws.Cells(rowcount, 11).Value = percent_change
                
                End If

                
                rowcount = rowcount + 1

                'Reset everything to let it run again
                total_vol = 0
                year_open = 0
                year_close = 0
                year_change = 0
                percent_change = 0
                
            End If
        Next i


    Next ws

End Sub

