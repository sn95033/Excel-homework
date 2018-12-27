
'------------------Excel VBA - Stock Price Exercise - Moderate -----------------------------------------------------

'Create a summary table for all the stocks on the page
    'Summarize the ticker,  the annual stock price change, the annual stock price % change, and the total volume for the year
    'Note this table is in columns I,J,K,L
        'Graphically indicate the annual stock price change:
        'Cells which have negative (-) gain should be colored red
        'Cells which have positive (+) gain should be colored green
        
'------------------Excel VBA - Stock Price Exercise - Hard -----------------------------------------------------
    
'Do the same calculation for all stocks (on all worksheets)
'Create an overall summary table for the stocks with the greatest annual % increase, % decrease, and total volume for the year
    'Note that this summary table is in Columns O,P,Q

'------------------Calculational Notes -------------------------------------------------------------------------------------
'The calculation of the annual Stock Price change and % stock price change is based on
'The opening day price of the first day of the year  and the closing day price of the last day of the year


Sub Stock_Summary():

    '--------------------------------------------------------------
    '  LOOP THROUGH ALL SHEETS
    '--------------------------------------------------------------

    For Each ws In Worksheets

    
        '----------------------------------------------------------------------------------------------------------------
        '       SETUP VARIABLES AND POINTERS FOR THE LASTROW OF THE RAW DATA AND THE SUMMARY TABLE CURRENT ROW
        '----------------------------------------------------------------------------------------------------------------
        
        'Create a variable to hold the stock name, the stock price at beginning and end of the year, the % change, and the total volume
        
        Dim Stock_Name As String
        
        Dim Yr_Open As Double
        Yr_Open = 0
        
        Dim Yr_Close As Double
        Yr_Close = 0
        
        Dim Stock_Price_Change As Single
        Stock_Price_Change = 0
        
        Dim Stock_Percent_Change As Single
        Stock_Percent_Change = 0
        
        Dim Vol_Total As Double
        Vol_Total = 0
        
       
        'Keep track of the row for each stock
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
        '-----------------------------------------------------------------------------------------------------------------
        '       SETUP HEADER IN THE SHEET
        '-----------------------------------------------------------------------------------------------------------------
        
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "% Change"
        ws.Cells(1, 12) = "Yearly Volume"
        ws.Cells(1, 15) = "Annual Results"
        ws.Cells(1, 16) = "Ticker"
        ws.Cells(1, 17) = "Value"
        ws.Cells(2, 15) = "Greatest %Gain"
        ws.Cells(3, 15) = "Greatest %Loss"
        ws.Cells(4, 15) = "Greatest Traded Volume"
        'ws.Cells(1, 13) = "Yr Open"
        'ws.Cells(1, 14) = "Yr Close"
        
        
        '----------------------------------------------------------------------------------------------------------------
        '       CREATE A SUMMARY TABLE FOR EACH STOCK WITH ITS TICKER SYMBOL, YEARLY CHANGE, %CHANGE, AND TRADED VOLUME
        '----------------------------------------------------------------------------------------------------------------
        
            For i = 2 To LastRow
            
                'Check to see if the current row is the first data for the ticker
                'If the ticker for the current and prior row don't match, then the current row is the first of the year
                'Then set the Yr_Open value
                
                
                If (ws.Cells(i - 1, 1) <> ws.Cells(i, 1)) Then
                    Yr_Open = ws.Cells(i, 3).Value
                        
                    'Add the Stock Volume total
                    'Vol_Total = Vol_Total + ws.Cells(i, 7).Value
                        
                'Check to see if the current row is the last data for the ticker
                'If the ticker for the current and next row don't match, then the current row is the last of the year
                'Then set the Yr_close value and load into the summary table
                        
                ElseIf (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                
                    'Set Stock Name
                    Stock_Name = ws.Cells(i, 1).Value
                    
                    'Add this data to the Traded volume for the year
                    Vol_Total = Vol_Total + ws.Cells(i, 7).Value
                    
                    'Set Stock Price End since the ticker is not the same
                    Yr_Close = ws.Cells(i, 6).Value
                    
                    'Detect if the Beginning Stock Price is zero
                    If (Yr_Open <> 0) Then
                        'Calculate the Stock change as:
                        'Opening price of the 1st day of the Year - Closing Price of the Last Day of the Year
                        Stock_Price_Change = Yr_Close - Yr_Open
                        Stock_Percent_Change = Stock_Price_Change / Yr_Open
                        
                    'Catch if the Yr_Open is zero and set the stock Price and % change to 0 to avoid overflow error
                    ElseIf (Yr_Open = 0) Then
                        Stock_Price_Change = 0#
                        Stock_Percent_Change = 0#
                        
                    End If
                    
                      'Check if this stock is the highest Gain%, highest Loss% or Max Trading volume and Update it
                    If (ws.Cells(2, 17).Value < Stock_Percent_Change) Then
                        ws.Cells(2, 17).Value = Stock_Percent_Change
                        ws.Cells(2, 16).Value = Stock_Name
                    End If
                        
                    If (ws.Cells(3, 17).Value >= Stock_Percent_Change) Then
                        ws.Cells(3, 17).Value = Stock_Percent_Change
                        ws.Cells(3, 16).Value = Stock_Name
                    End If
                    
                    If (ws.Cells(4, 17).Value < Vol_Total) Then
                        ws.Cells(4, 17).Value = Vol_Total
                        ws.Cells(4, 16).Value = Stock_Name
                    End If
                    
                        
                    '-------- Load the results into the table and making graphical changes
                    
                    'Print Stock Name in the Summary Table
                    ws.Range("I" & Summary_Table_Row).Value = Stock_Name
                    
                    'Print Stock Volume in the Summary Table
                    ws.Range("L" & Summary_Table_Row).Value = Vol_Total
                    
                    'Print Stock Annual Change in the Summary Table
                    'Color Green if there is an annual gain or Red if there is a loss
                    
                    ws.Range("J" & Summary_Table_Row).Value = Stock_Price_Change
                    If (ws.Cells(Summary_Table_Row, 10).Value < 0) Then
                        ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                    Else
                        ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                    End If
                    
                    'Print Stock Annual Percentage change in the Summary Table
                    ws.Range("K" & Summary_Table_Row).Value = Stock_Percent_Change
                    
                    'Print the Yr open and Yr close for debug purposes
                    'ws.Range("M" & Summary_Table_Row).Value = Yr_Open
                    'ws.Range("N" & Summary_Table_Row).Value = Yr_Close
                    
                    
                    '-------Housekeeping for the next iteration looking at the next ticker symbol
                    
                    'Increment to the next row in the summary table
                    Summary_Table_Row = Summary_Table_Row + 1
                    
                    
                    'Reset the stock total
                    Vol_Total = 0
                    
                    'Reset beginning of year stock price
                    Yr_Open = 0
                    
                    'Reset End of year Stock price
                    Yr_Close = 0
                    
        
            
                'In this case the ticker symbol is the same for the current, prior and next rows
                Else
                    'Sum up the Stock volume total
                    Vol_Total = Vol_Total + ws.Cells(i, 7).Value
                
                End If
            
            Next i
            
        '-----------------------------------------------------------------------------------------------
        '           CURRENT WORKSHEET COMPLETE
        '-----------------------------------------------------------------------------------------------
    
        Next ws
        
End Sub

