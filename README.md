# Stocks Analysis
## Overview of Project
In this project, I edited, or refactored, the solution code to loop through all the data one time in order to collect stock data to examine the entire stock market for 2017 and 2018. Then, I determined whether refactoring my code successfully made the VBA script run faster. 
## Results

**Original Code:**

```vb

Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    

    '1b) Create three output arrays
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
            
            

            '3d Increase the tickerIndex.
            
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

```

**Refactored Code:**

```vb

Sub AllStocksAnalysis()
   'Format the output sheet on All Stocks Analysis worksheet
        Worksheets("All Stocks Analysis").Activate
        Range("A1").Value = "All Stocks (2018)"
   'Create a header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"

   'Initialize array of all tickers
        Dim tickers(11) As String
            tickers(0) = "AY"
            tickers(1) = "CSIQ"
            tickers(2) = "DQ"
            tickers(3) = "ENPH"
            tickers(4) = "FSLR"
            tickers(5) = "HASI"
            tickers(6) = "JKS"
            tickers(7) = "RUN"
            tickers(8) = "SEDG"
            tickers(9) = "SPWR"
            tickers(10) = "TERP"
            tickers(11) = "VSLR"
   
   'Initialize variables for starting price and ending price
        Dim startingPrice As Single
        Dim endingPrice As Single
   
   'Activate data worksheet
        Worksheets("2018").Activate
    
    'Get the number of rows to loop over
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   'Loop through tickers
        For i = 0 To 11
            ticker = tickers(i)
            totalVolume = 0
    
    'loop through rows in the data
        Worksheets("2018").Activate
        For j = 2 To RowCount
    
    'Get total volume for current ticker
        If Cells(j, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(j, 8).Value
        End If
    
    'get starting price for current ticker
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               startingPrice = Cells(j, 6).Value
        End If

    'get ending price for current ticker
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               endingPrice = Cells(j, 6).Value
        End If
        Next j
       
    'Output data for current ticker
        Worksheets("All Stocks Analysis").Activatebean
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

        Next i
            Dim startTime As Single
            Dim endTime  As Single

            yearValue = InputBox("What year would you like to run the analysis on?")

        Next i
            startTime = Timer
            endTime = Timer
            MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
        
        Next i
End Sub


```

**Time on VBA_Challenge_2017**

![VBA_Challenge_2017]( https://github.com/CorinneBean/Stock-Analysis/blob/b3c2dc3356d3de2b475b41ce38bf7b7f6b4ed0e3/Resources/VBA_Challenge_2017.png)

**Time on VBA_Challenge_2018**

![VBA_Challenge_2018]( https://github.com/CorinneBean/Stock-Analysis/blob/b3c2dc3356d3de2b475b41ce38bf7b7f6b4ed0e3/Resources/VBA_Challenge_2018.png)

## Summary

- **What are the advantages or disadvantages of refactoring code?**
Refactoring code is used to restructure the existing body of code by altering its internal structure of the code, but it doesn't change the behavior of the code.*https://refactoring.com/*

	- Advantages
		- helps to optimize code so that future implimentations of improvements or features are easier to take place
		- Since code is refined and optimized it is easier to find and fix bugs
		- Reviewing and understanding code is easier since the code is see

	- Disadvantages
		- Is time consuming if you are working with large code. As a result one should not attempt to refactor code if they have a deadline approaching. 
		- If not careful you can actually introduce bugs during refactoring
		- It is expensive and risky in the view of management. *https://www.c-sharpcorner.com/article/pros-and-cons-of-code-refactoring/*
	
- **How do these pros and cons apply to refactoring the original VBA script*
	- Refactoring the original VBA script made it more efficient. Since it was a small code sample the refactoring took minimal time. In the end the advantages out weighed the disadvantages since it was a fairly quick and easy project.