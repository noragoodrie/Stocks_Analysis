# Stocks_Analysis

**Project Overview**

This week I had to take a dive into the world of VBA to get Steve the information on Green Stocks that he needs for his parent’s investment! Visual Basic Application (VBA) is the main way I was able to compare and grab the right data for Steve this week. VBA gives the ability to generate codes into your spreadsheet in a timely manner as well as incorporating interactive boxes to get all the data you need onto 1 worksheet!

**Results**

Below you can see 2 charts that were created for all stocks in 2017 and another for all stocks. In 2018. Each chart has 3 columns which include a ticker, total daily volume, and a return. In the return column you can see either an increase or a decrease on the stock accounts. The data that is filled in green shows that there was an increase over the year, and the data. Filled in red shows a decrease within the year. 

<img width="252" alt="Screen Shot 2021-10-10 at 8 34 03 PM" src="https://user-images.githubusercontent.com/91299616/136735062-bf757528-d6e6-4f4b-8683-188f28e2f061.png">

<img width="252" alt="Screen Shot 2021-10-10 at 8 34 18 PM" src="https://user-images.githubusercontent.com/91299616/136735068-5f8170cb-2117-45a4-a05b-5ec44cfff585.png">

**Analysis**

The first code I that I ran to get these results was long and strainious. For the worksheet “All Stocks Analysis” I had to generate 3 separate Sub’s which included, sub AllStocksAnalysis, Sub formatingAllStocksAnalysis, and Sub ClearFormat. Having to run all 3 of these subs to get the results I needed took approximately .6895 for both 2017 and 2018 stock analysis. 
We then had to refactor our code by incorporating the instructions file to our VBA and getting all of the data from the 3 subs onto 1 sheet. After filling in the blanks myself I was able to get my codes run time down to approximately .1887 seconds which as we can see if a HUGE decrease in time! Below you can see the generated code with the steps used to grab our data.

    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
    
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
        
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
            End If
        
        '3c) 'If  Then
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If

            '3d Increase the tickerIndex.
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            
            End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        

**Summary**

Some advantages of refactoring code are the fact that you can make your code look more organized and cleaner right off the bat. Instead of having all this code created on a handful of different Subs you are able to have one Sub … to End Sub with ALL the information you need right there. I believe it gives you the ability to not have as many mistakes in your work as well. 

A disadvantage to refactoring you might not have the time to refactor your code. If you are pressed on a project your time may be limited so its important to do it the right way in the beginning. Also, refactoring code doesn’t mean changing your inputs or loops it just means to reorganize your data into small chunks and loops.

After refactoring the All Stocks Analysis I found that my code ran faster than the original VBA script due to the organization on the same Sub sheet instead of having it broken down in 3 different subs! Below is the 2017 All Stock Analysis Run time compared to the 2017 All Stock Analysis Refactored run time.

<img width="254" alt="VBA_Challenge_2017 Before" src="https://user-images.githubusercontent.com/91299616/136735342-c94cfebd-b656-4d40-84ce-b1c6643ad8a7.png">


<img width="252" alt="Screen Shot 2021-10-10 at 9 56 35 PM" src="https://user-images.githubusercontent.com/91299616/136735301-3287c8d3-c8c0-470d-932b-2a88418860ea.png">


