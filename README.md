# Stock-Analysis
Analysis of Green energy stocks

## Table of Contents
- [Overview of Project](#OverviewProject)
  * [Background](#Background)
  * [Purpose](#purpose)
- [Initial Analysis](#Analysis)
- [Results](#results)
- [Summary](#summary)
- [References](#references)


## <a name="OverviewProject"></a>Overview of Project
### <a name="Background"></a>Background
Steve, needs to analyze the performance of green energy stocks (from the year 2017 and 2018) for his parents. They beleive that as the fossil fuel is being used up, there will be more reliance on alternative energy production. Steve's parents are particularly interested in a green energy stock named "Daqo New energy Corporation (DQ)", which makes silicon wafers for Solar panels. He further needs to help them diversify their investment in other such category stocks. For this purpose I am using Microsoft <i><b>Visual Basic for application (VBA) </b> </i> to create a workbook for Steve to analyze the stocks <i><b> yearly return </b></i> and <i><b>total daily volume</b></i> at the click of a button.
 
### <a name="Purpose"></a>Purpose
Now that Steve is able to analyze a few stocks successfully, he wants to expand the dataset to include the entire stock market over the last few years. Although our code works well for a dozen stocks, it might not work very efficiently for thousands of stocks. That is, it might take a long time to execute to display the required analysis.
Since the data is huge, we would want our macro to run efficiently. For this purpose we will <b><i>refactor</b></i> our existing code to improve clarity and the run time performance. This project highlights the importance of <b><i>refactoring</i></b> by comparing the run time for our previous code and the new refactored code. Please refer to the [References](#references) section to view both the codes.  


## Original Code 
 - **Ticker** </br>
      In the original code we have stored the tickers in an array. We have hardcoded the values of the 12 tickers for which we need to run the stock analysis. The declaration of this array is shown below:
      ```
      Dim tickers(11) As String
      ```
          
 - **Nested For Loops** </br>
     We run the first loop to read each ticker from the array.
     ```
     For i = 0 To 11
       ticker = tickers(i)
          ---- all the code goes in here
     Next i
     ```
     For each ticker read, we activate the sheet for the year selected by the user and the run a **nested loop** inside the main loop to read all the records for that ticker.
     ```
     For j = 2 To RowCount
     --- The code explained below goes here
     Next j
     ```
     We then calculate the **total volume** for that ticker and store in the variable **"totalVolume"** based on the below condition:
    ```
    If Cells(j, 1).Value = ticker Then
       totalVolume = totalVolume + Cells(j, 8).Value
    End If
     ```
     Next we get the **stating price** of the ticker in the following way:
      ```
      If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If
      ```
     Calculate the **ending price** of the ticker in the following way:
  ```
     If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        endingPrice = Cells(j, 6).Value
     End If
   ```        
  Next, in the main loop we again activate the other sheet and populate the Ticker value, Total volume and Yearly return  values.
  ```
   Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
  ```
 ## Refactored Code
 In the refactored code we have declared 3 arrays to store the values of "total volume", "Starting Price","Ending Price" of all the tickers.
 ```
 'Create three output arrays  
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
 ```
 We create an index "tickerindex" to access the arrays declared above.
 ```
 'Create a ticker Index
  tickerIndex = 0
 ```
 We loop over all the rows in our worksheet as shown below:
 ```
For i = 2 To RowCount
   --- The rest of the code explained below goes here
Next i
```
'Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        'Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         If Cells(i, 1).Value <> " " And Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
         End If
            
            
            
        'End If
        
        'check if the current row is the last row with the selected ticker
         'If the next rows ticker doesnt match, increase the tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            tickerIndex = tickerIndex + 1
        End If
        
 ## <a name="Analysis"></a>Analysis
### <a name="allAnalysis"></a>All stock Analysis for 2017 and 2018
The table below displays the analysis of a dozen green stocks for the years 2017 and 2018. The table mainly consists of 3 columns:
1. Ticker
2. Total Daily Volume
3. Return
<p align = "center">
<img src = "Module2_VBA\Images\2017_data.png" width = 25%>
<img src = "Module2_VBA\Images\2018_data.png" width = 25%>
</p>        
     
#### **Total Volume and Yearly Return:**
Total Volume is an indicator of the popularity of the stock. With the help of the bar graphs below we see that the DQ stock was traded 3 times higher in 2018 in comaprison to the previous year 2017.  Hence even though DQ was much popular in 2018, yet it resulted in a loss, with an annual return of -63%. 
Comparing the total volume of the dozen stocks in both the years, we see that all the stocks had a higher traded volume in 2018. But the annual return for almost all the stocks has resulted in a loss in 2018, except for 2 stocks "ENPH","RUN" which has resulted in a gain. Also, "SPWR" has the highest total volume for both the years 2017 and 2018.

Since the majority of the stocks have a negative yearly return, this might also be because the stock market may have gone down during the year 2018. Even though by looking at the data in the year 2018, DQ doesnt seem to be a good stock to invest in, there might be other reasons for its decline. Hence we need to analyze more data of the previous years for this stock to see its yearly trends.

<p align = "center">
<img src = "Module2_VBA\Images\2017_volbar.png" width = 25%>
<img src = "Module2_VBA\Images\2018_volbar.png" width = 25%>
</p>

#### **Execution time of Original Codes:**
The execution times of the original code for the years 2017 and 2018 is as below: 
<p align = "center">
!<img src = "Module2_VBA\Images\2017_runtime.png1" width = 50%>
!<img src = "Module2_VBA\Images\2018_runtime.png1" width = 50%>
</p>

## <a name="results"></a>Results
Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

## <a name="summary"></a>Summary
In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?

## <a name="references"></a>References
[1] [Original Code](green_stocks.xlsm)
[2] [Refactored Code](VBA_Challenge.xlsm)
