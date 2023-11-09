Sub StockData():

Dim ws As Worksheet    'Declaring a variable for as worksheet to use in For Loop

    For Each ws In ThisWorkbook.Worksheets                 'For Loop to run the same code through all worksheets available
                
    Dim I As Long                      'declaring row iterator
    Dim row_count As Long              'declaring for total number of rows
    Dim totalVol As LongLong           'declaring for total volume traded
    Dim summary_row As Long            'declaring for rows to be populated
    Dim counter As Long                'declaring for number of times a ticker appeared
    Dim summary_row_count As Long      'delcaring for total rows of summary
    Dim greatest_gain As Double        'declaring for greatest % gain
    Dim greatest_loss As Double        'declaring for greatest % loss
    Dim greatest_totalVol As LongLong  'declaring for greatest total Volume
        
    row_count = ws.Cells(Rows.Count, "A").End(xlUp).Row   'retrieving total number of rows
        
    'Assigining Titles
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    summary_row = 2 'Summarized data will be shown from second row hence value is 2
    
        'Looping through the Data to perform tasks
        For I = 2 To row_count      'Loop will start comparing data from second row hence value is 2
            
            If (ws.Cells(I, 1).Value = ws.Cells(I + 1, 1).Value) Then      'if statement compare value of a cell to cell in next row
                totalVol = totalVol + ws.Cells(I, 7).Value              'retrieve the total Volume by adding all the values which belongs to same ticker
                counter = counter + 1                                'showing how many times a specific ticker shows up in data set
                            
            Else                                                     'if Value in next row is not same, else block will be executed
                totalVol = totalVol + Cells(I, 7).Value              'when value of next cell is not same as the cell we are in,total volume will be concluded and stored in summary row
                counter = counter + 1
                
                'Output the data
                ws.Cells(summary_row, 9).Value = ws.Cells(I, 1).Value
                ws.Cells(summary_row, 10).Value = ws.Cells(I, 6).Value - ws.Cells(I - counter + 1, 3).Value
                ws.Cells(summary_row, 11).Value = FormatPercent(ws.Cells(summary_row, 10).Value / ws.Cells(I - counter + 1, 3).Value)
                ws.Cells(summary_row, 12).Value = totalVol
           
                counter = 0                                      'counter set to zero start counting again for next ticker
                totalVol = 0                                     'variable set to zero to calculate total volume for next ticker
                summary_row = summary_row + 1                    'Now data will be collected for the next summary_row
            
            End If
               
            If ((ws.Cells(I, 10).Value > 0) And (ws.Cells(I, 11).Value > 0)) Then
                      
                ws.Cells(I, 10).Interior.ColorIndex = 4
                ws.Cells(I, 11).Interior.ColorIndex = 4
                
            ElseIf ((ws.Cells(I, 10).Value < 0) And (ws.Cells(I, 11).Value < 0)) Then
                      
                ws.Cells(I, 10).Interior.ColorIndex = 3
                ws.Cells(I, 11).Interior.ColorIndex = 3
                
            End If
            
        Next I
        
        'With statement allows you to perform a series of statements on a specified object without requalifying the name of the object.
        With ws.Range("J2:K" & row_count).FormatConditions.Add(xlCellValue, xlGreater, 0)
        
            With .Interior
            
                .ColorIndex = 4
            
            End With
            
        End With
        
        With ws.Range("J2:K" & row_count).FormatConditions.Add(xlCellValue, xlLess, 0)
        
            With .Interior
            
                .ColorIndex = 3
            
            End With
        
        End With
            
              
        ws.Range("O2").Value = "Greatest % increase"            'Labeling the Rows for corresponding Values
        ws.Range("O3").Value = "Greatest % decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        summary_row_count = ws.Cells(Rows.Count, "K").End(xlUp).Row  'retrieving number of rows in summary data
        
        greatest_gain = Application.WorksheetFunction.Max(ws.Range("K2:K" & summary_row_count))       'Worksheet Function to get Max Value in myRange1 i.e. Column K
        greatest_loss = Application.WorksheetFunction.Min(ws.Range("K2:K" & summary_row_count))       'Worksheet Function to get Min Value in myRange1 i.e. Column K
        greatest_totalVol = Application.WorksheetFunction.Max(ws.Range("L2:L" & summary_row_count)) 'Worksheet Function to get Max Value in myRange2 i.e. Column L
        
        'Output the Data
        ws.Range("Q2").Value = FormatPercent(greatest_gain)
        ws.Range("Q3").Value = FormatPercent(greatest_loss)
        ws.Range("Q4").Value = greatest_totalVol
        
        Dim row_greatest_gain As Long             'Declaring a variable to get row in which greatest_gain exist
        Dim row_greatest_loss As Long             'Declaring a variable to get row in which greatest_loss exist
        Dim row_greatest_totalVol As Long         'Declaring a variable to get row in which greatest_totalVol exist
                
        row_greatest_gain = Application.WorksheetFunction.Match(greatest_gain, ws.Range("K:K"), 0)  'retrieving position of row where greatest gain exist using Worksheet Function Match
        ticker_greatest_gain = ws.Cells(row_greatest_gain, 9).Value     'retrieving ticker of greatest gain
        ws.Range("P2").Value = ticker_greatest_gain                     'Output the Data
        
        row_greatest_loss = Application.WorksheetFunction.Match(greatest_loss, ws.Range("K:K"), 0)  'retrieving position of row where greatest loss exist using Worksheet Function Match
        ticker_greatest_loss = ws.Cells(row_greatest_loss, 9).Value     'retrieving ticker of greatest loss
        ws.Range("P3").Value = ticker_greatest_loss                      'Output the Data
        
        
        row_greatest_totalVol = Application.WorksheetFunction.Match(greatest_totalVol, ws.Range("L:L"), 0)  'retrieving position of row where greatest total volume exist using Worksheet Function Match
        ticker_greatest_TotalVol = ws.Cells(row_greatest_totalVol, 9).Value     'retrieving ticker of greatest total Volume
        ws.Range("P4").Value = ticker_greatest_TotalVol                         'Output the Data
        
        ws.Columns("I:Q").AutoFit         'Fit the Data into the cells
          
    Next ws
    
End Sub



