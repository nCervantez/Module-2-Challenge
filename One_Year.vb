Sub Stock_Cleanup()

        'Setting the necessary variables for this script
        Dim TSVolume As Double
        Dim StockTableRow As Integer
        Dim Ticker As String
        Dim YearlyChange As Integer
        Dim Percent As Integer
        Dim YearOpen As Integer
        Dim YearClose As Integer
        Dim YearChange As Integer
        
        
        
        'This will be called on later in the code. Starts the row at 2 to avoid the column Headers
        StockTableRow = 2
        
        'This line of code will determine the last row for the script to run on
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'These lines will insert the proper headers into the new columns
        Range("I1").Value = "Ticker"
        
        Range("J1").Value = "Yearly Change"
        
        Range("K1").Value = "Percent Change"
        
        Range("L1").Value = "Total Stock Volume"
        
        'For loop that will start at row 2 and end at the last determined row
        For i = 2 To LastRow
            
            'This If statement will check if the next cell is of the same ticker or not
            'if the next ticker is different It will Add the current ticker,
            'and total stock volume to the correct rows
            'and will reset to begin counting for the next ticker
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                'This will calculate the change in the yearl open and yearly
                'close values for the ticker
                YearClose = Cells(i, 6).Value
                
                YearChange = YearClose - YearOpen
                
                Ticker = Cells(i, 1).Value
                
                TSVolume = TSVolume + Cells(i, 7).Value
                
                Range("I" & StockTableRow).Value = Ticker
                
                Range("L" & StockTableRow).Value = TSVolume
                
                Range("J" & StockTableRow).Value = YearChange
                
                Range("K" & StockTableRow).Value = (YearChange / YearOpen) * 100
                
                'Will reset the variables to be used for the next ticker
                StockTableRow = StockTableRow + 1
                
                YearOpen = 0
                
                YearClose = 0
                
                YearChange = 0
                
                TSVolume = 0
            
            Else
            
                'Will add the daily stock volume if the tickers are the same
                'and will capture the year open volume to be used in
                'calculations later
                TSVolume = TSVolume + Cells(i, 7).Value
                YearOpen = Cells(i, 3).Value
            
            End If
            
        Next i
        
    
End Sub

