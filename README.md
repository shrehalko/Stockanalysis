# Stock-Analysis
Analysis of Green energy stocks

## Table of Contents
- [Overview of Project](#OverviewProject)
  * [Background](#Background)
  * [Purpose](#purpose)
- [Pseudo Code](#pseudocode)
  * [Original Code](#Original)
  * [Refactored Code](#refactored)
- [Results](#results)
  * [Total Volume and Yearly Return](#TVYR)
  * [Execution Time](#time)
- [Summary](#summary)
  * [Advantages or Disadvantages of Refactoring code](#advantages)
  * [How do these Pros and Cons apply to refactoring the original VBA script?](#pros)
- [References](#references)


## <a name="OverviewProject"></a>Overview of Project
### <a name="Background"></a>Background
Steve, needs to analyze the performance of green energy stocks (from the year 2017 and 2018) for his parents. They beleive that as the fossil fuel is being used up, there will be more reliance on alternative energy production. Steve's parents are particularly interested in a green energy stock named "Daqo New energy Corporation (DQ)", which makes silicon wafers for Solar panels. He further needs to help them diversify their investment in other such category stocks. For this purpose I am using Microsoft <i><b>Visual Basic for application (VBA) </b> </i> to create a workbook for Steve to analyze the stocks <i><b> yearly return </b></i> and <i><b>total daily volume</b></i> at the click of a button.
 
### <a name="Purpose"></a>Purpose
Now that Steve is able to analyze a few stocks successfully, he wants to expand the dataset to include the entire stock market over the last few years. Although our code works well for a dozen stocks, it might not work very efficiently for thousands of stocks. That is, it might take a long time to execute to display the required analysis.
Since the data is huge, we would want our macro to run efficiently. For this purpose we will <b><i>refactor</b></i> our existing code to improve clarity and the run time performance. This project highlights the importance of <b><i>refactoring</i></b> by comparing the run time for our previous code and the new refactored code. Please refer to the [References](#references) section to view both the codes.  

## <a name="pseudocode"></a>Pseudo Code
### <a name="Original"></a>Original Code 
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
 ### <a name="refactored"></a>Refactored Code
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
        
## <a name="results"></a>Results

### <a name="allAnalysis"></a>All stock Analysis for 2017 and 2018

### <a name="TVYR"></a>Total Volume and Yearly Return

The table below displays the analysis of a dozen green stocks for the years 2017 and 2018. The table mainly consists of 3 columns:
1. Ticker
2. Total Daily Volume
3. Return
<p align = "center">
<img src = "Module2_VBA\Images\2017_data.png" width = 30%>
<img src = "Module2_VBA\Images\2018_data.png" width = 30%>
</p>        
     
**Total Volume** is an indicator of the popularity of the stock. From the above images we see that the **DQ stock** was traded **3 times higher in 2018 in comaprison to the previous year 2017.**  Hence even though **DQ** was much popular in 2018, yet it **resulted in a loss, with an annual return of ***-63%***.** <br> 
Comparing the **total volume** of the dozen stocks in both the years, we see that majority of the stocks had a higher traded volume in **2018**. But the annual return for almost all the stocks has resulted in a loss in 2018, except for 2 stocks **"ENPH","RUN"** which has resulted in a gain. <br> Also, **"SPWR"** has the highest total volume for both the years **2017 and 2018**.

Since the majority of the stocks have a negative yearly return, this might also be because the stock market may have gone down during the year 2018. <br> **Even though by looking at the data in the year 2018, DQ doesn't seem to be a good stock to invest in, there might be other reasons for its decline. Hence we need to analyze more data of the previous years for this stock to see its yearly trends.**


### <a name="time"></a>**Execution time:**

The execution times of the ***Original*** code for the years **2017 and 2018** is as below: 
<p align = "center">
<img src = "Module2_VBA\Images\2017_runtime.png" width = 30%>
<img src = "Module2_VBA\Images\2018_runtime.png" width = 30%>
</p>

The execution times of the ***Refactored*** code for the years **2017 and 2018** is as below: 
<p align = "center">
<img src = "Module2_VBA\Images\2017_refactor.png" width = 30%>
<img src = "Module2_VBA\Images\2018_refactor.png" width = 30%>
</p>

**From the above images we conclude that the ***refactored code*** executed much faster.** 

## <a name="summary"></a>Summary
### <a name="advantages"></a>Advantages or Disadvantages of Refactoring code

Improving the design of the existing code has many advantages:
1. **Efficiency:** Refactoring increases the efficiency of the code. Our code executes much faster.
2. **Readability:** A refactored code increases the readability and the understanding of the complex code. 
3. **Helps finding Bugs**: Since a refactored complex code is more readable and easir to understand, it helps the coders to find and fix bugs quickly. 
4. **Saves Time:** Code refactoring reduces the likelihood of errors and helps the future coders to implement new functionality easily.
5. **Reduced Cost:** A clean and well-structured code takes less to update and maintain.  

The major **disadvantage of refactoring** is:
1. **Time consuming:** Refactoring can be a time consuming process, especially when the coder does not understand the functionality of the code.
2. **Risky:** If the application is big,refactoring might be risky,as it should not hamper the existing functionality of the code. 
3. **Introduce more bugs:** If the coder does not have enoough time to understand the functionality of the code, refactoring might introduce more bugs which can make the system unstable. 

## <a name="pros"></a>How do these Pros and Cons apply to refactoring the original VBA script?
1. Refactoring our original VBA script, helped our code execution to be much faster than before.
2. By introducing arrays to store the 'total volume','starting price' and 'ending price' we have increased the readability of our code. Also we have removed the redundant code from our original script. In future if we have more stocks to analyze, we can just increase the array size and run our code without any modifications.
3. The **major challenge** faced by me in refactoring the code was that it was **time consuming**. Being new to VBA I had to spend more time understanding the syntaxes and various functionalities and how to implement them in my code.

## <a name="references"></a>References
[1] [Original Code](Module2_VBA/green_stocks.xlsm) <br> 
[2] [Refactored Code](Module2_VBA\Challenge\VBA_Challenge.xlsm.xlsm)
