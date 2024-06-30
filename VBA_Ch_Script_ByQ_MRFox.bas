Attribute VB_Name = "Module1"
Sub VBAChallengeByQuarter()

    '=======================================================================================================
    'SCRIPT A LOOP THAT OUTPUTS EACH TICKER AND THE QUARTERLY CHANGE, PERCENT CHANGE, AND TOTAL STOCK VOLUME
    '=======================================================================================================
    
    'Set up the loop and define the last row
    Dim row As Long
    Dim column As Integer
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, "A").End(xlUp).row
    
    'Script stores values of
    Dim ticker As String
    Dim volume_total As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim price_change As Double
    Dim percent_change As Double
    
    'Outline the Summary Table
    Dim summary_table_row As Integer
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Quarterly Change"
    Cells(1, 12).Value = "Percent Change"
    Cells(1, 13).Value = "Total Volume"
    
    'Set initial values
    volume_total = 0
    open_price = Cells(2, 3).Value
    price_change = 0
    summary_table_row = 2
    
    'Loop through the rows
    For row = 2 To lastrow
        
        'Check if the next row is the same and if not
        If Cells(row + 1, 1).Value <> Cells(row, 1).Value Then
        
            'Set the ticker value and print it to the summary table
            ticker = Cells(row, 1).Value
            Range("J" & summary_table_row).Value = ticker
            
            'Add to total volume, print it to the summary table, and reset value to 0
            volume_total = volume_total + Cells(row, 7).Value
            Range("M" & summary_table_row).Value = volume_total
            volume_total = 0
            
           'Get the close price
            close_price = Cells(row, 6).Value
            
            'Calculate price change and print to summary table
            price_change = close_price - open_price
            Range("K" & summary_table_row).Value = price_change
            
            'Calculate the percent change and print to summary table
            percent_change = price_change / open_price
            Range("L" & summary_table_row).Value = percent_change
            
            'Calculate the new open price based on the next row
            open_price = Cells(row + 1, 3).Value
                
            'Add one to the summary table row
            summary_table_row = summary_table_row + 1

        'When the next row is the same
        Else
            
            'Add to the volume total
            volume_total = volume_total + Cells(row, 7).Value

        End If

    Next row

    '========================================================================================
    'ADD FUNCTIONALITY FOR GREATEST % INCREASE, GREATEST 5 DECREASE AND GREATEST TOTAL VOLUME
    '========================================================================================

    'Define the last row of summary table
    Dim lastrowsumtbl As Long
    lastrowsumtbl = Cells(Rows.Count, "J").End(xlUp).row

    'Script stores values of
    Dim inc_ticker As String
    Dim great_perc_inc As Double
    Dim dec_ticker As String
    Dim great_perc_dec As Double
    Dim tvol_ticker As String
    Dim great_tot_vol As Double
    
    'Outline the Summary Table
    Cells(1, 17).Value = "Ticker"
    Cells(1, 18).Value = "Value"
    Cells(2, 16).Value = "Greatest Percent Increase"
    Cells(3, 16).Value = "Greatest Percent Decrease"
    Cells(4, 16).Value = "Greatest Total Volume"
    
    'Set initial values
    great_perc_inc = -1
    great_perc_dec = 1
    great_tot_vol = 0
    
    'Loop through the rows
    For row = 2 To lastrowsumtbl
        
        'Check if the next row is the same and if not
        If Cells(row, 12).Value > great_perc_inc Then
        
            great_perc_inc = Cells(row, 12).Value
            inc_ticker = Cells(row, 10).Value
        
        ElseIf Cells(row, 12).Value < great_perc_dec Then
        
            great_perc_dec = Cells(row, 12).Value
            dec_ticker = Cells(row, 10).Value
        
        End If
        
        If Cells(row, 13).Value > great_tot_vol Then
        
            great_tot_vol = Cells(row, 13).Value
            tvol_ticker = Cells(row, 10).Value
        
        End If
        
        'Print the greatest percent increase
        Range("Q2").Value = inc_ticker
        Range("R2").Value = great_perc_inc
        Range("Q3").Value = dec_ticker
        Range("R3").Value = great_perc_dec
        Range("Q4").Value = tvol_ticker
        Range("R4").Value = great_tot_vol

        'Format new table
        Range("R2").NumberFormat = "0.00%"
        Range("R3").NumberFormat = "0.00%"
        
    Next row

    '============================
    'APPLY CONDITIONAL FORMATTING
    '============================

    'Loop through the rows
    For row = 2 To lastrow

        'If the value is negative then format as red
        If Cells(row, 11).Value < 0 Then
            Cells(row, 11).Interior.ColorIndex = 3

        'If the value is positive then format as green
        ElseIf Cells(row, 11).Value > 0 Then
            Cells(row, 11).Interior.ColorIndex = 4
                    
        End If

    Next row

    'Format columns and autofit
    Columns(11).NumberFormat = "0.00"
    Columns(12).NumberFormat = "0.00%"
    Cells.EntireColumn.AutoFit
    
    'Confirm completion of script
    MsgBox "Complete"

End Sub

