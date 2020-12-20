Attribute VB_Name = "Module1"
Option Explicit

Sub ColumnNames()
'This subroutine adds the column headers for unique stock analysis

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 14).Value = "Maximum Yearly Percent Change"
    Cells(3, 14).Value = "Minimum Yearly Percent Change"
    Cells(4, 14).Value = "Largest Stock Volume"
    Cells(1, 15).Value = "Stock"
    Cells(1, 16).Value = "Quantity"

End Sub

Sub StockAnalysis()
'This sub routine finds the yearly open price, close price, change in yearly price and percent change in price _
by looping through the entire data set searching for specific conditions. _
The sub also uses multiple if statements to run calculation sbased on certain criteria.
    
    Dim LastRowFull As Double
    'Last row of full data set
    Dim UniqueCounter As Double
    'Counter used to move down the unique summary data columns
    
    Dim OpenPrice As Double
    'Firs open price of the year for each stock
    Dim ClosePrice As Double
    'Close price at end of the year for each stock
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim StockVolume As Double
    
    Dim i As Double
    
    Dim MaxValue As Double
    'Holds the larges running percent change
    Dim MaxRow As Double
    'Finds the row count of the maximum percent change
    
    Dim MinValue As Double
    'Holds the smallest (largest negative) running percent change
    Dim MinRow As Double
    'Finds the row with the smallest (most negative)ent change
    
    Dim MaxVol As Double
    'Holds thergest stock volume
    Dim MaxVolRow As Double
    'Finds the row count of the maximum stock volume
    
    LastRowFull = Cells(Rows.Count, 1).End(xlUp).Row
    'Calculates last row of the full stock data sey for the year
    
    UniqueCounter = 2
    
    StockVolume = 0
    
    For i = 2 To LastRowFull
    'External for loop to loop through the unique stock symbol data set
        
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        'If current stock symbol is not the same as the row above, then consider it a unique stock symbol
            Cells(UniqueCounter, 9).Value = Cells(i, 1).Value
            OpenPrice = Cells(i, 3).Value
        End If
        
        If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
        'If the stock symbol in the current row is the same as the next one then aggregate stock volume
            StockVolume = StockVolume + Cells(i, 7).Value
        End If
        
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        'Determines close price and total stock volume by finding the last instance of a repeating stock symbol within the range and _
        designates it as row for the close price
    
            ClosePrice = Cells(i, 6).Value
            
            YearlyChange = ClosePrice - OpenPrice
            
            StockVolume = StockVolume + Cells(i, 7).Value
            'aggregate of stock volume
            Cells(UniqueCounter, 10).Value = YearlyChange
            
            Cells(UniqueCounter, 12).Value = StockVolume
            
            If OpenPrice = 0 Then
            'Use to helpiminate the division byro error cause when OpenPrice=0
                PercentChange = 0
                Else
                    PercentChange = YearlyChange / OpenPrice
            End If
            
            Cells(UniqueCounter, 11).Value = PercentChange
                         
            StockVolume = 0
            'reset stock volume to zero after finding the close price, meaning that the last continuous stock symbol has ended
            
            UniqueCounter = UniqueCounter + 1
            'increase counter by one, meaning move down one row on the unique stock symbol after the last repetative stock symbol has been found
        End If
       
    Next i
    
End Sub

Sub FormatUnique()
'This sub uses interior color formatting for percentages less than 0 and greater than 0, formats dolar values and percent values as well as auto fits the columns for the data
    Dim LastRowUnique As Double
    Dim i As Double
    
    LastRowUnique = Cells(Rows.Count, 11).End(xlUp).Row
    'finds last unique row within summary data set
    
    For i = 2 To LastRowUnique
        
        If Cells(i, 11).Value > 0 Then
            Cells(i, 11).Interior.ColorIndex = 4
        ElseIf Cells(i, 11).Value < 0 Then
            Cells(i, 11).Interior.ColorIndex = 3
        End If
        
    Next i
    
    Range("J2:J" & LastRowUnique).NumberFormat = "$#,###0.00"
    Range("C:F").NumberFormat = "$#,###0.00"
    Range("K2:K" & LastRowUnique).NumberFormat = "0.00%"
    Cells(2, 16).NumberFormat = "0.00%"
    Cells(3, 16).NumberFormat = "0.00%"
    Columns("A:P").AutoFit

End Sub

Sub Bonus()
'Bonus uses a for loop to calculate the maximum percent stock change, minimum percent change and maximum stock volumes
    Dim i As Double
    
    Dim MaxValue As Double
    'Holds the larges running percent change
    Dim MaxRow As Double
    'Finds the row count of the maximum percent change
    
    Dim MinValue As Double
    'Holds the smallest (largest negative) running percent change
    Dim MinRow As Double
    'Finds the row with the smallest (most negative)ent change
    
    Dim MaxVol As Double
    'Holds thergest stock volume
    Dim MaxVolRow As Double
    'Finds the row count of the maximum stock volume
    
    Dim LastRowUnique As Double
    
    LastRowUnique = Cells(Rows.Count, 9).End(xlUp).Row
    
    MaxValue = -100
    
    MinValue = 100
    
    MaxVol = 1
    
    For i = 2 To LastRowUnique
    
        If Cells(i, 11).Value > MaxValue Then
        'Finds the maximumtock perecent change as the sub loops through the data
            MaxValue = Cells(i, 11).Value
            MaxRow = i
        End If
                
        If Cells(i, 11).Value < MinValue Then
        'Finds the minimum tock perecent change as the sub loops through the data
            MinValue = Cells(i, 11).Value
            MinRow = i
        End If
                
        If Cells(i, 12).Value > MaxVol Then
        'Finds the largest total stock volume
            MaxVol = Cells(i, 12).Value
            MaxVolRow = i
        End If
    
    Next i
    
    Cells(2, 15) = Cells(MaxRow, 9).Value
    Cells(2, 16) = MaxValue
    
    Cells(3, 15) = Cells(MinRow, 9).Value
    Cells(3, 16) = MinValue
    
    Cells(4, 15) = Cells(MaxVolRow, 9).Value
    Cells(4, 16) = MaxVol
    
End Sub

Sub MultiSheetFullProcess()
'This subroutine calls all of the individual sub routines and compiles them into a single code
    
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ' The above lines of code help make the code run more efficiently by turning off auto calcs and screen updates
    
    Dim WS_Count As Integer
    Dim i As Integer

    Dim StartTime As Double
    Dim SecondsElapsed As Double
    
    StartTime = Timer
    'timer used to measure how long the code takes to execute
    
    WS_Count = ActiveWorkbook.Worksheets.Count
    'finds the total number of worksheets within the workbook for loop
    
    For i = 1 To WS_Count
    'For i = 1 To 1
        
        Sheets(i).Activate
        
        Call ColumnNames
        
        Call StockAnalysis
        
        Call Bonus
        
        Call FormatUnique
    
    Next i
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
   
    SecondsElapsed = Round(Timer - StartTime, 2)
    
    MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation
    
    
End Sub

