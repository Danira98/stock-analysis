# ***Stock-Analysis with VBA***

## Overview of Project:

### Backgroud:

In this project, we aid Steve by analyzing twelve green energy stocks, including DAQO New Energy Corporation with the hopes to find a trend to see which stocks will be ideal for his parents to invest in aside from DAQO. We focus our analysis on the years of 2017 and 2018,and return the total daily volume and the yearly return percentage of the year 2018 for all companies. To expedite the analysis , we used VBA with Excel and with the help of a button, we carried our analysis at a faster rate. We carry this analysis with two different methods, one with a standard VBScript and another with a refactored VBSript.

### Purpose:

The purpose of this project is to create a vba script that will return the total daily volume and the yearly return percentage at a faster rater(refactored vba script), with the worksheet formated in a way that is easier to read, and with designated colors in the return column to see which stocks are ideal to invest in. The buttons added to our worksheet will make it easier for Steve to run and clear the worksheet without having to go back and forth to the Macros option.

## Results:

### Code:

The AllStocksAnalysisRefactored code was designed to take less time to analyse the data but in order to accomplish this, we had to modify our initial AllStocksAnalysis code.

The first change we made was create three arrays to hold the ticker volumes, the starting prices and the ending prices, and a ticker index to use as a variable in our for loops. We use these new arrays that hold 12 variables in the for loops we created.
```
'1a) Create a ticker Index
    
    Dim tickerIndex As Single 
    
            tickerIndex = 0

    '1b) Create three output arrays
    
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
```

Instead of using nested loops like the original code, we use for loops and use the tickerIndex variable as the index to find our values. We create our first for loop to initialize our tickerVolumes variables equal to zero. Next, we create a for loop to find the starting price and ending price.It runs the same way as the previous code but without the nested loop and we increase tickerIndex by 1 at the end to keep runing our loop.
```
For i = 0 To 11

        tickerVolumes(i) = 0
        
    Next i
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
                     tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            

            '3d Increase the tickerIndex
                '(the previous if statement checks already if the next ticker doesn't match the previous row's ticker)
            
                tickerIndex = tickerIndex + 1
            
        End If
    
    Next i

```
We also include a new for loop to return the value of our arrays tota volumes, ending price and starting price
```
'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
```
We also included our formating at the end of our code instead of creating a new VBScript to do it.

### Results:

After running our VBScripts, We are able to obain the following results:

![Worksheet ss](https://user-images.githubusercontent.com/111034667/188967961-7be383e0-17fd-4a2b-baec-76a3e84989fd.png)

We can observe that the companies with the positive return are ENPH and RUN, and if we pay attention to the company DAQO, abbreviated as DQ, it has a -62.6% which means that it would be ideal for Steve  to advice his parents to consider investing in stocks like ENPH and RUN instead or other companies with a smaller negative percentage retunr like VSLR or TERP if they choose to invest in multiple stocks.

By running both VBScripts, AllStocksAnalysis and AllStocksAnalysisRefactored, we can observe that the refactored code runs faster with the same outcome.The following images show the elapsed times for each code and the worksheet outcome.

![AllStocks](https://user-images.githubusercontent.com/111034667/188970079-0efeebed-69f0-4beb-86b3-c2fbe3619a82.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/111034667/188970102-89498d8d-a2c5-44d7-8775-d675036e751e.png)

<sub>AllStocksAnalysis results and elapsed time <sub>
  
  ![AllStocksR](https://user-images.githubusercontent.com/111034667/188970275-d5a3fe47-6ef6-4c23-9ed2-2e8e7fbc0d2a.png)
![VBA_Challenge_2018 (2)](https://user-images.githubusercontent.com/111034667/188970295-53868707-6872-4156-b61e-afb2d487cc22.png)

  <sub>AllStocksAnalysisRefactored results and elapsed time <sub>
    
 ## Summary:
    
 ### Advantages and Disadvantages of Refactored code:
 Coding in any programming language can be a tedious task, and our goal as a programmer is to make a code that runs efficiently enough to get to our goal and as clear  as time allows us to. By refactoring our code, we will be able to use less memory, improve the logic of the code which makes it easier to read to future uses,and it usually takes fewer steps. Unfortunately, although refactoring our code is ideal to make it more efficient, often times programmers do not have the time to spend thinking about how to make their code more efficient;they usually create a code that will be able to get us to the goal within a deadline given, which sometimes can lead to codes that take up more memory but that work perfectly fine for our end goal.
    
### Advantages and Disadvantages of Refactored VBScript:
    The refactored VBScript allows us to run our analysis faster, and overall it is easier to understand , even with beginner knowledge of coding VBA. It unfortunately does take longer to code since we have to think of ways to make the code easier to understand and more efficient. I personally had trouble understanding how some of the loops were going to be set up but after going back and forth with the original script, reading the hints given in the assignment, and following a similar structure to the hint,  I was able to figure out a code that worked. The advantage of our original script is that it was easier to set up,smaller,smaller amount of loops needed since nested loops were used, but it did take longer to run and also the fomatting of our worksheet wasn't included, which lead us to create and additional VBScript to do so, whereas our refactored code includes the formatting.
    
