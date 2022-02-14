# Stocks_Analysis
module 2 Excel &amp; VBA

Module 2 project focused on further developing skills with Excel. In this module we learned about core concepts in coding with the language VBA. 
During this module we learned how to write and run code for subroutines. This allows us to automate tasks within Excel.
Overall, VBA has allowed me to search across multiple stocks at a time, extrapolate data regarding stock tickers, daily volumes, return, and has allowed me to visualize this information
in Excel with the user by creating a simple button to press to run the macro. The purpose of this challenge was to enforce concepts such as initializing variables and arrays, activating relevant worksheets,
running 'for' loops with conditional 'if' statements to give us output values such as each individual ticker, the total daily volume, and return for each stock in the years 2017/8.
Information is then visualized with the "All Stocks Analysis" page to not only show our outputs, but to display our outputs in a manner that is easy to read and understand so that quick decisions can be made with each respective stock
    Results of the 12 stocks has given us some insight. Due to the wonderful year of returns of "DQ" during 2017 (almost 200%!), our collegue was interested in buying "DQ" stock.
However if we analyze the year of 2018, it is clear "DQ" had some struggles, down 62% from last year. What is quite pecuiliar, is volume from 2017 to 2018 went up 300%! From the simple results from the chart,
it is clear as the total daily volume increased for "DQ", the return for the stock decreased. Massive buys for "DQ" stock took place during 2018 (no doubt thanks to the returns of 2017).
Leading us to conlude as more people bought "DQ" stock, the demand of the stock soon diminished, therefore decreasing price. At this time, if "DQ" volume continues to increase, it is not advisable to buy this stock as it will
likely be accompanied by furthere return decrease. On the contrary, "ENPH and "RUN" both saw good returns both years. It can be seen as their volume increases, so do their returns.
At this time if volume continues to increase it can be associated with an increase in returns. "ENPH" and "RUN" stocks have high probability of good returns. Overall, the majority of our listed stocks saw
decline in the year 2018 due to other outside factors.
    Original VBA script "All Stocks Analysis" included a nested for loop with 3 conditional statements with another output command to display data output:

    For t = 0 To 11
    ticker = tickers(t)
    totalVolume = 0
    
    Worksheets("2018").Activate
    
    For p = 2 To RowCount
    If Cells(p, 1).Value = ticker Then
    totalVolume = totalVolume + Cells(p, 8).Value
    
    End If
    
    If Cells(p - 1, 1).Value <> ticker And Cells(p, 1).Value = ticker Then
        
        startingPrice = Cells(p, 6).Value
        
    End If
    
    If Cells(p + 1, 1).Value <> ticker And Cells(p, 1).Value = ticker Then
    
        endingPrice = Cells(p, 6).Value
        
    End If
    
Next p
    
Worksheets("All Stocks Analysis").Activate
Cells(4 + t, 1).Value = ticker
Cells(4 + t, 2).Value = totalVolume
Cells(4 + t, 3).Value = endingPrice / startingPrice - 1

in contrast, the refactored code is similar, but is written as 3 separate for loops with a couple of conditionals

For i = 0 To 11
        tickerVolumes(i) = 0
    
    Next i
    
 
For j = 2 To RowCount
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value

         If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(j, 3).Value 'column 3 is our 'open' column in the spreadsheet
            
        End If
            
            
      
          If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value 'column 6 is our 'closing' column in the spreadsheet
            
          tickerIndex = tickerIndex + 1
          
        End If
        
    
    Next j
    


    Worksheets("All Stocks Analysis").Activate
   
    
    For tdv = 0 To 11
       Cells(tdv + 4, 2).Value = tickerVolumes(tdv) 
        Cells(tdv + 4, 3).Value = tickerEndingPrices(tdv) / tickerStartingPrices(tdv) - 1
        Cells(tdv + 4, 1) = tickers(tdv)
        
    Next tdv



Advantages of refactored code include faster run times, more concise code containing 'for' loops with arrays, and consolidation of all the subroutines used in the "All Stocks Analysis" into one subroutine.
Advantages to keeping the original code: nested 'for' loops are powerful, one array was required, subroutines can be individualized. Disadvantages include: much longer run time, must format output one subroutine at a time
and the largest con is refactoring can be time consuming.
