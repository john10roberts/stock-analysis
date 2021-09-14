# stock-analysis
Stock data analysis
## Overview of Project
Refactor previous version of stock analysis visual basic for applications code to improve efficiency. 

### Purpose
The purpose of this project was to take a previously created Excel Macro Enabled workbook that was used to gather stock information for a client and improve the existing VBA code to improve the efficieny of the workbook.    

## Analysis and Challenges
I began by taking the existing code and removing the existing loop structure while leaving the code that dealt with the formatting.  I then began working through the looping structure to improve the efficiency of the previously written code.  The steps I went through for the refactoring are as follows:

1a) Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = 0

1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i

2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount

3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If

3c) check if the current row is the last row with the selected ticker. 
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If

3d Increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If

 4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11

        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
    Next i



### Results
![Original 2017 Speed](https://github.com/john10roberts/kickstarter-analysis/blob/main/Resources/Theater_Outcomes_vs_Launch.png)
![Original 2018 Speed](https://github.com/john10roberts/kickstarter-analysis/blob/main/Resources/Theater_Outcomes_vs_Launch.png)
![Refactored 2017 Speed](https://github.com/john10roberts/kickstarter-analysis/blob/main/Resources/Theater_Outcomes_vs_Launch.png)
![Refactored Speed](https://github.com/john10roberts/kickstarter-analysis/blob/main/Resources/Theater_Outcomes_vs_Launch.png)
For the analysis based on launch date we used the year column to allow us to filter the data based on the year the kickstarter was started.  Created a pivot table to show all the successfull, failed and canceled kickstarters.  Then we filtered that by the parent category of theater.  That data showed us again that most successful kickstarters for theaters are at their peak in May.  The successful campaigns start to trail off the following months and hit a low in december.  

### Analysis of Outcomes Based on Goals
![Outcomes Based on Goals](https://github.com/john10roberts/kickstarter-analysis/blob/main/Resources/Outcomes_vs_Goals.png)
For the outcomes based on goals we created a table that would calculate the outcomes of a particular kickstarter based on a range of goals.  The ranges started at under 1000 - greater than 50000.  We then used these groupings to calculate the percentage of kickstarters that suceeded/failed/cancelled.  The data shows that lower goal kickstarters are successful at a much higher rate than.  And while the line chart associated with the data shows that there was an increase in the percentage successful for campaigns from 35-45k the total number of those campaigns are so low that it might be a little misleading just looking at the chart.  It appears that for kickstarters with goals of less than 5k they have the highest odds of success.  

### Challenges and Difficulties Encountered
There didn't appear to be any difficulties or challenges with this data set.  Everything was presented in a logical manner.

## Results

- What are two conclusions you can draw about the Outcomes based on Launch Date?
Kickstarters launched in May have the best chance of success.  
Kickstarters launched in December have the smallest chance of success

- What can you conclude about the Outcomes based on Goals?
Kickstarters of less than 5k are the most successful
Kickstarters greater than 5k are the least successful

- What are some limitations of this dataset?
We don't know much information about the kickstarters other than a name and blurb, more information on what it was would be beneficial.  It would also be nice to be able to see the type of donor and maybe some information on the donors to be able to classify the type of person donating. 

- What are some other possible tables and/or graphs that we could create?
Backers to success it would be nice to know if there is any correlation between the number of backers and a projects success. 
Backers Descriptive Statistics 
