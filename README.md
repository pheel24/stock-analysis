# Stock Analysis

## Overview
This project is an extension of our coursework in this module in which we analyzed a few particular stocks with VBA analysis. In this extension rather then looking at only a few stocks we retooled our analysis to handle many stock performances over a particular year. In addition to expanding the scope of our VBA script, we were hoping to decrease the runtime of the script. 

# Results
Most of the additions to the module code were for loops, intended to expand the scope of the script and allow it to handle many different kinds of stocks. The script is written such that stock data with the same formatting can be fed into it, rather then the script being tied to a few different stocks within our data. 
First we created a null ticker index and set up 3 null arrays to hold the output, like so:

'Create a ticker Index
    
    tickerIndex = 0

'Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
Our first for loop intialized the ticker volume to zero, while a simulataneous for loop covered all the rows of the data. Within the second loop we created logic that populates the null arrays created above based on certain conditions with respect to the position in the data (e.g. "if current row is the last row with selected ticker then increase ticker index")

'Create a for loop to initialize the tickerVolumes to zero.
    
    For tickerIndex = 0 To 11
       tickerVolumes(tickerIndex) = 0
        
 'Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount
    
        'Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        End If
        
        'Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                    
        End If

        
        'check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
         If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6)
                    
         End If
            

            'Increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                   tickerIndex = tickerIndex + 1
            
            
        End If
    
    Next i
    
We made one more for loop to loop through the created arrays to output our metrics of interest: Ticker, Total Daily Volume, and Return.

'Loop through your arrays to output the Ticker, Total Daily Volume, and Return
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + tickerIndex, 1).Value = tickers(tickerIndex)
        Cells(4 + tickerIndex, 2).Value = tickerVolumes(tickerIndex)
        Cells(4 + tickerIndex, 3).Value = (tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex)) - 1
        
    Next tickerIndex
    
In both cases however, our code unfortunately did not run faster. Though for 2018 the times were quite close.

Before the refactor for 2017:

![2017_no_refactor](https://user-images.githubusercontent.com/95315957/154413482-0e6fb4d4-acf6-4aca-8369-6c71c48faf2a.PNG)

After:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/95315957/154413567-13daee00-0fc7-48e8-9740-97e129cfd40a.PNG)

Before the refactor for 2018:

![2018_no_refactor](https://user-images.githubusercontent.com/95315957/154413657-a6a7cc8c-6c12-458d-b06f-24e3b531adbb.PNG)

After:

![VBA_Challenge_2018](https://user-images.githubusercontent.com/95315957/154413670-b8561d54-492d-4e11-a235-3b9dbc3359ec.PNG)

This is more than likely due to the additional logic we wrote into this script, demonstrating a tradeoff between scope and speed. I would argue the tradeoff is worth it in this case, as the times are not very different from eachother. Where our first script runs faster it is less applicable to other stocks without editing, which is the strong point of our refactored script. 

In general, greater efficiency is always worth striving for. However, given that most models have more parameters than the speed of analysis to consider it seems worth it to exchange some speed for applicability. 
