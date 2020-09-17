'Stock Market Analysis VBA Script Written By Carter Alvarez

Sub StockMarketAnalysis()

    'Loop through all of the worksheets
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    

        'Column Headers
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"

        'Declare Variables
        Dim Ticker_Symbol As String
        Dim Ticker_Stock_Volume As Double
        Dim Last_Row As Long
        Dim Summary_Table_Row As Integer
        Dim Yearly_Open As Double
        Dim Yearly_Close As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
            Ticker_Stock_Volume = 0
            Yearly_Change = 0
            Summary_Table_Row = 2
            Yearly_Open = ws.Cells(2, 3).Value

        'To Find the Last Row
        Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'For Loop to Fill the Desired Stock Information in our New Locations
        For i = 2 To Last_Row

        'Put a check to make sure we group the same ticker names together, if the next row is not the same then...
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

            'Identify TickerSymbol
            Ticker_Symbol = ws.Cells(i, 1).Value
            
            'Add to Ticker Volume
            Ticker_Stock_Volume = Ticker_Stock_Volume + ws.Cells(i, 7).Value

            'Print the Ticker Symbol in our New Ticker Name Column
            ws.Range("I" + CStr(Summary_Table_Row)).Value = Ticker_Symbol
            
            'Print the Ticker Total Stock Volume In our Summary Table
            ws.Range("L" + CStr(Summary_Table_Row)).Value = Ticker_Stock_Volume
            
            'Set the Yearly_Open, Yearly_Close, and Yearly_Change Values
            Yearly_Close = ws.Cells(i, 6).Value
            
            Yearly_Change = (Yearly_Close - Yearly_Open)
            
            'Print Yearly_Change
            ws.Range("J" + CStr(Summary_Table_Row)).Value = Yearly_Change
            
            'Set the Percent Change while Accounting for it being 0
            If Yearly_Open = 0 Then
            Percent_Change = 0
            
            Else
            Percent_Change = (Yearly_Change / Yearly_Open)
            End If
            
            'Format and Print the Column to use % and Decimals
            ws.Range("K" + CStr(Summary_Table_Row)).NumberFormat = "0.00%"
            ws.Range("K" + CStr(Summary_Table_Row)).Value = Percent_Change

            
            'Add to the Summary_Table_Row
            Summary_Table_Row = Summary_Table_Row + 1
            
            'Reset Ticker_Stock_Volume Total to 0 for the next Ticker
            Ticker_Stock_Volume = 0
            
            'Reset Yearly_Open
            Yearly_Open = ws.Cells(i + 1, 3)
       
            
            Else
            
              'Add to the Stock Volume Total
              Ticker_Stock_Volume = Ticker_Stock_Volume + ws.Cells(i, 7).Value
            
            End If
        

        Next i
        
   
    
           
    'Format to Highlight Green for a Positive Change and Red for a Negative Change
    LastRow_Summary = ws.Cells(Rows.Count, 9).End(xlUp).Row

       For i = 2 To LastRow_Summary
               If Cells(i, 10).Value >= 0 Then
                   Cells(i, 10).Interior.ColorIndex = 4
               Else
                   Cells(i, 10).Interior.ColorIndex = 3
               End If
       Next i
    
    Next ws

End Sub
