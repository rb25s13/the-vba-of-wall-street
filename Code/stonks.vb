Sub stonks()

'    Create a script that will loop through all the stocks for one year and output the following information:
'
'    1.) The ticker symbol.
'
'    2.) Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'
'    3.) The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'
'    The total stock volume of the stock.
'
'    You should also have conditional formatting that will highlight positive change in green and negative change in red.
    
    Dim ws As Worksheet
    'Set ws = ActiveSheet
    
    'loop through all sheets
    For Each ws In ActiveWorkbook.Worksheets
        
        ws.Select
        
        With ws
        
            Application.ScreenUpdating = False

            Dim column  As Integer
            Dim tickercolumn As Integer
            Dim yearlychange As Integer
            Dim percentchange As Integer
            Dim totalvolume As Integer
            Dim sumrow  As Long
            Dim stockvol As Long
            Dim tickerfirstrow As Double
            Dim totalstockvolume As Long
            Dim sumrange As Range
            
            'format the columns to auto-fit
            Cells.Columns.AutoFit
            
            'determine the last row in the ticker column
            lastrow = Cells(Rows.Count, 1).End(xlUp).Row
            
            'variables created to hold column and row values
            column = 1
            tickercolumn = 9
            yearlychange = 10
            percentchange = 11
            totalvolume = 12
            sumrow = 2
            stockvol = 7
            tickerfirstrow = 2
            
            'write the headers and format them
            Range("I1").Value = "Ticker"
            Range("J1").Value = "Yearly Change"
            Range("K1").Value = "Percent Change"
            Range("L1").Value = "Total Stock Volume"
            Range("P1").Value = "Ticker"
            Range("Q1").Value = "Value"
            Range("O2").Value = "Greatest % Increase:"
            Range("O3").Value = "Greatest % Decrease:"
            Range("O4").Value = "Greatest Total Volume:"
            Range("A1", "Q1").Font.Bold = True
            Range("A1", "Q1").HorizontalAlignment = xlCenter
            Range("P1", "Q4").HorizontalAlignment = xlCenter
            Range("O2", "O4").Font.Bold = True
            Range("O2", "O4").HorizontalAlignment = xlRight
            
            'loop through columns to find ticker symbol groups
            For currow = tickerfirstrow To lastrow
                
                'if current row does not match the second row, note the ticker symbol group
                If Cells(currow + 1, column).Value <> Cells(currow, column).Value Then
                    
                    'write ticker symbol name to I column
                    Cells(sumrow, tickercolumn).Value = Cells(currow, column).Value
                    
                    'write yearly change by subtracting first day open value from final day close value
                    Cells(sumrow, yearlychange).Value = Cells(currow, column + 4).Value - Cells(tickerfirstrow, column + 3)
                    
                    'read yearly change value to determine whether the value is positive
                    If Cells(sumrow, yearlychange).Value > 0 Then
                        
                        'if the cell is positive, change the fill to green
                        Cells(sumrow, yearlychange).Interior.ColorIndex = 4
                        
                        'read yearly change value to determine whether the value is negative
                    ElseIf Cells(sumrow, yearlychange).Value < 0 Then
                        
                        'if the cell is negative, change the fill to red
                        Cells(sumrow, yearlychange).Interior.ColorIndex = 3
                        
                    End If
                    
                    'read the yearly change value to determine if the value is equal to 0
                    If Cells(sumrow, yearlychange).Value = 0 Then
                        
                        'if the value equals 0, then write the percent change as 0
                        Cells(sumrow, yearlychange).Value = 0
                        
                        'read the yearly change value to make sure the value is not 0
                    ElseIf Cells(sumrow, yearlychange).Value <> 0 And Cells(tickerfirstrow, column + 2) <> 0 Then
                        
                        'divide the yearly change value by the first day open value
                        percentage = Cells(sumrow, yearlychange).Value / Cells(tickerfirstrow, column + 2)
                        
                    End If
                    
                    'format percent change value to show as percentage
                    Cells(sumrow, percentchange).Value = FormatPercent(percentage, 2)
                    
                    'add all of the values of stock volume for the ticker symbol range
                    Cells(sumrow, totalvolume).Value = Application.Sum(Range(Cells(tickerfirstrow, stockvol), Cells(currow, stockvol)))
                    
                    'increment the summary row for the next stock ticker
                    sumrow = sumrow + 1
                    
                    'update the value of the first row for the ticker symbol by adding 1 to the current row
                    tickerfirstrow = currow + 1
                    
                End If
                
            Next currow
            
            'BONUS - return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
            
            'determine the last row in the summary table column
            sumlastrow = Cells(Rows.Count, 11).End(xlUp).Row
            
            'setting the range of the table of total percent change
            Set sumrange = Range(Cells(2, percentchange), Cells(sumlastrow, 11))
            'setting the range of the  table of total stock volume
            Set sumvolrange = Range(Cells(2, totalvolume), Cells(sumlastrow, 12))
            
            'name the max value in the range of percent change
            greatestincrease = Application.WorksheetFunction.Max(sumrange)
            
            'name the min value in the range of percent change
            greatestdecrease = Application.WorksheetFunction.Min(sumrange)
            
            'name the max value in the range of total stock volume
            moststockvolume = Application.WorksheetFunction.Max(sumvolrange)
            
            'name the row of the greatest increase value
            girow = WorksheetFunction.Match(greatestincrease, sumrange, 0) + sumrange.Row - 1
            
            'name the row of the greatest decrease value
            gdrow = WorksheetFunction.Match(greatestdecrease, sumrange, 0) + sumrange.Row - 1
            
            'name the row of the most stock volume value
            msvrow = WorksheetFunction.Match(moststockvolume, sumvolrange, 0) + sumvolrange.Row - 1
            
            'write the greatest stock increase ticker
            Cells(2, 16).Value = Cells(girow, 9)
            
            'write the greatest stock decrease ticker
            Cells(3, 16).Value = Cells(gdrow, 9)
            
            'write the most stock volume ticker
            Cells(4, 16).Value = Cells(msvrow, 9)
            
            'write the greatest stock increase value
            Cells(2, 17).Value = greatestincrease
            
            'write the greastest stock decrease value
            Cells(3, 17).Value = greatestdecrease
            
            'write the most stock volume value
            Cells(4, 17).Value = moststockvolume
            
            'format the greatest stock increase value to percentage
            Cells(2, 17).Value = FormatPercent(Cells(2, 17).Value, 2)
            
            'format the greatest stock decrease value to percentage
            Cells(3, 17).Value = FormatPercent(Cells(3, 17).Value, 2)
            
            'format the columns to auto-fit
            'Cells.Columns.AutoFit
            
        End With
        
    Next ws
    
    Application.ScreenUpdating = True
    
End Sub
