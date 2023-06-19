Sub Stock_Cleanup()

        'Setting the necessary variables for this script
        Dim TSVolume As Double
        Dim StockTableRow As Integer
        Dim Ticker As String
        Dim YearlyChange As Double
        Dim YearOpen As Double
        Dim YearClose As Double
        Dim YearChange As Double
        Dim openprice As Boolean
        
        
        'Variable to have the for loop grab the first iteration of the ticker
        openprice = False
        
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
                    
                    'Conditional formatting to color cells based on gain or loss on yearly change
                    If Range("J" & StockTableRow).Value < 0 Then
                        
                        Range("J" & StockTableRow).Interior.ColorIndex = 3
                    
                    Else
                    
                        Range("J" & StockTableRow).Interior.ColorIndex = 4
                        
                    End If
                    
                Range("K" & StockTableRow).Value = (YearChange / YearOpen)
                
                'Will reset the variables to be used for the next ticker
                StockTableRow = StockTableRow + 1
                
                YearOpen = 0
                
                YearClose = 0
                
                YearChange = 0
                
                TSVolume = 0
                
                'reset the openprice variable to allow for next ticker to capture year open price
                openprice = False
            
            Else
            
                'Will add the daily stock volume if the tickers are the same
                'and will capture the year open volume to be used in
                'calculations later
                TSVolume = TSVolume + Cells(i, 7).Value
                'This will grab the open price in each loop then will not function until the next ticker is being used
                If Not openprice Then
                    
                    YearOpen = Cells(i, 3).Value
                    openprice = True
                
                End If
                
            
            End If
            
        Next i
        
        
        
        'These lines of code will find the greatest percent increase and decrease
'-----------------------------------------------------------------------------------
        LastRowPercent = Cells(Rows.Count, 11).End(xlUp).Row
        LastRowStock = Cells(Rows.Count, 12).End(xlUp).Row
        Dim PercentR As Range
        Dim VolR As Range
        Dim Max As Double
        Dim Min As Double
        Dim MaxVolume As Double
        'Added seperate tickers to test
        Dim Ticker2 As String
        Dim Ticker3 As String
        Dim Ticker4 As String
        Set PercentR = Range("K2", "K" & LastRowPercent)
        Set VolR = Range("L2", "L" & LastRowStock)
        'Setting the greatest table
        
        
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        'Excel worksheet functions that will find the highest and lowest values in my set range
        
        Max = Application.WorksheetFunction.Max(PercentR)
        MsgBox (Max)
        Min = Application.WorksheetFunction.Min(PercentR)
        MsgBox (Min)
        MaxVolume = Application.WorksheetFunction.Max(VolR)
        MsgBox (MaxVolume)
            
        For j = 2 To LastRowPercent
            
            'If statement only needs to capture the max value in the range
            'Max value already found in the range
            If Cells(j, 10).Value <> Max Then
                
            
            Else
                
                Range("Q2").Value = Max
                Ticker2 = Cells(j, 9).Value
                
                Range("P2").Value = Ticker2
                
            End If
            
        Next j
        
        For k = 2 To LastRowPercent
            
            'If statement only needs to capture the min value in the range
            'min value already found in the range
            If Cells(k, 10).Value <> Min Then
                
            
            Else
                
                Range("Q3").Value = Min
                Ticker3 = Cells(k, 9).Value
                
                Range("P3").Value = Ticker3
                
            End If
            
        Next k
        
        For l = 2 To LastRowStock
            
            'If statement only needs to capture the max value in the range
            'Max Stock Volume already found in the range
            If Cells(l, 12).Value <> MaxVolume Then
                
            
            Else
                
                Range("Q4").Value = MaxVolume
                Ticker4 = Cells(l, 9).Value
                
                Range("P4").Value = Ticker4

            End If

        Next l
End Sub
