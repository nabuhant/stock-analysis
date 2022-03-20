# Stock Analysis (VBA)
Module 2 Challenge


## Overview of Project
### Background
Steve's parents want to invest in the stock market, specifically the Daqo stock. Steve has stock market data for a set of stocks for the years 2017 and 2018. He wants to analyze the data to be able to advise his parents on their investment. He wants to advise them based on the total volume (shares traded throughout the year, measures stock activity) and yearly return (percentage difference in price from the beginning of the year to the end of the year). These analyses are to be carried out using VBA, a programming language that allows tasks’ automation within Excel. 

### Purpose
During the module the original VBA script was implemented such that the calculation of total volume and return of each stock is carried out based on the year the user enters, then the worksheet with the output (analysis) is formatted based on set conditions and properties. The challenge's purpose is to refactor the script such that it runs more efficiently therefore in a shorter time while performing the same functionality.

## Results
### Script Runtime Comparison (Original vs. Refactored)
The table below demonstrates the run time for both scripts (original script written throughout the module and refactored script). For both 2017 and 2018 stock data, the refactored script had a shorter runtime in comparison to the original script. 

|      | 2017 | 2018 |
| :--- |:----:| :---:|
| **Original Script**   | ![image1](/Resources/AllStockAnalysis_2017.png) | ![image3](/Resources/AllStockAnalysis_2018.png) |
| **Refactored Script** | ![image2](/Resources/VBA_Challenge_2017.PNG)    | ![image4](/Resources/VBA_Challenge_2018.PNG)    |

#### Some factors that resulted in the decrease of the runtime when comparing the original and refactored script include:  
**1.** The elimination of the nested *For loop* as in the code snippets below.

In the original script, there are two nested *For loops* such that the code runs 36,156 times (12XRowCount). The outside loop increases the ticker index, while the inner loop carries out the *If statements* used to find the total volume and values required to calculate the yearly return. 
```
'Original Script
    For i = 0 To 11
        
        totalVolume = 0
        ticker = tickers(i)
        Worksheets("2018").Activate 'READ
        
        '5) loop through rows in the data
        For j = 2 To RowCount
```
In the refactored script, the ticker index is a declared variable incremented by 1 in the final *If statement* that looks for the last row of the current ticker.
```
'Refactored Script
    tickerIndex = 0 'Declared Variable
    
         If Cells(i, 1).Value = ticker And Cells(i + 1, 1).Value <> ticker Then ' If statement that results in index incrementation
            ...
            
            '3d Increase the tickerIndex since it's the last row of current ticker
            tickerIndex = tickerIndex + 1
         End If

```
**2.** The elimination of the *If statement* below from the original script below. It is not necessary in the refactored script since the ticker index is incremented when the last row of the current index is reached (occurs in the refactored script snippet above).
```
'Original Script
           If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
           End If
```
### Stock Performance Comparison (2017 vs. 2018)
The table below demonstrates the stock performance output of both the original and refactored script. As seen below both scripts gave the exact same output.

|      | Original Script | Refactored Script |
| :--- |:----:| :---:|
| **2017** | ![image5](/Resources/AllStockAnalysisResults_2017.png) | ![image6](/Resources/VBA_ChallengeResults_2017.PNG) |
| **2018** | ![image7](/Resources/AllStockAnalysisResults_2018.png) | ![image8](/Resources/VBA_ChallengeResults_2018.PNG) |  

For 2017, 14 out of the 15 stocks being analyzed had a return over zero with 4 stocks having a return over 100%. For 2018, only 2 stocks had a return over zero, with the lowest return being -62.6%. Overall, the performance in 2017 was better.  
  
Based on the stock performance, especially the return value, the stocks advised for investment are **ENPH** and **RUN**. **ENPH** demonstrated an increase in its return of 129.5% in 2017, which dropped to 81.9% but is still a relatively high return. **RUN** demonstrated an increase in its return of 5.5% in 2017, which went to increase to a return of 84.0% in 2018 (a huge improvement in comparison to 2017).

## Summary
#### The General Advantages/Disadvantages of Refactoring Code
An advantage of refactoring code is the production of a code that is more flexible. A more flexible code allows for the addition of new functions easily, without jeopardizing the functionality of the pre-established code. Another advantage of refactoring code is the production of a code that is more efficient. An efficient code is more reliable, requires less resources to execute, and uses relatively easy to understand/follow logic.  
A disadvantage of refactoring code, especially more complicated code, is the time it requires.
#### Advantages/Disadvantages of Refactoring Applied to the Original VBA Script
In this challenge, the refactored script is more efficient such that it runs in a shorter time compared to the original script. 
