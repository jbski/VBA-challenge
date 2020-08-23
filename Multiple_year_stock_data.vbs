Sub stock_analysis()

Dim ticker, max_tv_ticker, max_pd_ticker, max_pi_ticker As String
Dim beg_open, end_close As Double
Dim stock_volume, max_pctincrease, max_pctdecrease As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim tot_volume, max_totalvolume As Double
Dim ws As Worksheet
Dim lrow, last_sum_ticker_row As Long
Dim i, j, k As Long



For Each ws In Worksheets

    
    'Clear the values and formats of the given range
    ws.Range("I:Q").ClearContents
    ws.Range("I:Q").ClearFormats

    'set the last row in each worksheet
    lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Populate the header rows
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'initialize the variable j so the correct rows will be used in
    'the calculated fields in columns I:L
    j = 1
    
    'volume and the open price on Jan 01 are also set here to be used
    'in the total volumen and change calculations
    stock_volume = ws.Cells(2, 7).Value
    beg_open = ws.Cells(2, 3).Value

    'loop through each row of stock data to find the total volume, opening price
    'and closing price to be used in the analysis in columns ("I:Q")
    For i = 2 To lrow

        ' increment the value of stock_volume for each ticker
        ' the total will then be used for the summary ticker in "I"
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
            stock_volume = stock_volume + ws.Cells(i + 1, 7).Value
        End If

        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            end_close = ws.Cells(i, 6).Value

            yearly_change = end_close - beg_open
            
            'check if opening price is zero to avoid divide by zero errors.
            If beg_open = 0 Then
                percent_change = 0
            Else
                percent_change = ((end_close - beg_open) / beg_open)
            End If
            

            j = j + 1
            ws.Cells(j, 9).Value = ticker
            ws.Cells(j, 10).Value = yearly_change
            ws.Cells(j, 11).Value = percent_change
            ws.Cells(j, 12).Value = stock_volume
                        
            'assign the beginning open price and stock volume for the new ticker
            beg_open = ws.Cells(i + 1, 3).Value
            stock_volume = ws.Cells(i + 1, 7).Value

            'set the cell colors to red if yearly change is negative
            'and green if the yearly change is positive
            If yearly_change < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 4
            End If

         End If
         
         

    Next i

    'apply the correct formatting to each column
    ws.Range("J2:J" & lrow).NumberFormat = "#,##0.00"
    ws.Range("K2:K" & lrow).NumberFormat = "#,##0.00%"
    ws.Range("L2:L" & lrow).NumberFormat = "###0"

    'find the last row of the ticker values in column "I"
    last_sum_ticker_row = Cells(Rows.Count, 9).End(xlUp).Row


    max_pctincrease = o
    max_pctdecrease = o
    max_totalvolume = o


    'loop through the summarized ticker data to find
    'the greatest % increase, the greatest % decrease and
    'the greatest stock volume
    For k = 2 To last_sum_ticker_row

        If ws.Cells(k, 11).Value > max_pctincrease Then
            max_pctincrease = ws.Cells(k, 11).Value
            max_pi_ticker = ws.Cells(k, 9).Value
        End If

        If ws.Cells(k, 11).Value < max_pctdecrease Then
            max_pctdecrease = ws.Cells(k, 11).Value
            max_pd_ticker = ws.Cells(k, 9).Value
        End If

        If ws.Cells(k, 12).Value > max_totalvolume Then
            max_totalvolume = ws.Cells(k, 12).Value
            max_tv_ticker = ws.Cells(k, 9).Value
        End If

    Next k

    'Populate the row labels
    ws.Range("O1").Value = "Greatest % Increase"
    ws.Range("O2").Value = "Greatest % Decrease"
    ws.Range("O3").Value = "Greatest Total Volume"

    'Populate the ticker values for the greatest increase,
    'decrease and total volume
    ws.Range("P1").Value = max_pi_ticker
    ws.Range("P2").Value = max_pd_ticker
    ws.Range("P3").Value = max_tv_ticker

    'populate the value for the greatest % increase, greatest % decrease,
    'and greatest total volume
    ws.Range("Q1").Value = max_pctincrease
    ws.Range("Q2").Value = max_pctdecrease
    ws.Range("Q3").Value = max_totalvolume

    'Apply formatting to the % values and the total volume
    ws.Range("Q1:Q2").NumberFormat = "###0.00%"
'    ws.Range("Q3").NumberFormat = "#,##0"
    ws.Range("Q3").NumberFormat = "0.0000E+00"

    'Format the column widths
    ws.Columns("I:L").AutoFit
    ws.Columns("O:Q").AutoFit


Next ws


End Sub

