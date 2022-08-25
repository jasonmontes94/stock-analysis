### Stock Analysis

## Overview

A friend of mine,Steve, wanted to analyze a set of stocks from various companies from the years 2017 and 2018. The data he provided had enough information to gather the yearly return (the difference between the stock price at the beginning of the year to the end of the year), however we needed to provide code that calculated it for us.

Initially, we had several subroutines that created new headers, ran code that showed us the yearly return for specific stocks, and then formatted the information for us to easily see if there was a negative or positive return. We also included a timer to see how long the code took to run. We had then refactored the code to include every task in one single subroutine, called AllStocksAnalysisRefactored().

## Results

The first difference between the original code we used to calculate yearly return for the stocks was to include a ticker index in the refactored code:

```
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
```

We then had to make sure to loop through all the rows in the spreadsheet data after initializing the tickerVolumes to zero:

```
'2a- Create a For loop to initialize the tickerVolumes to zero
For i = 0 To 11

    tickerVolumes(i) = 0
        
Next i

'2b- Loop over all the rows in the spreadsheet
For i = 2 To RowCount

'3a- Increase volume for current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
```

The main difference here, compared to before we refactored the code, is that we are now looping every row in the spreadsheet instead of the specific stocks we were looking for. How would this affect the time it took to run the code? Let’s take a look.

Here is the time it took for each year before we refactored the code.
For 2017:

![2017_initial_timer.png](/Resources/2017_initial_timer.png)

For 2018:

![2018_initial_timer.png](/Resources/2018_initial_timer.png)

Now, here is the timer for the refactored code.

For 2017:

![VBA_Challenge_2017.png](/Resources/VBA_Challenge_2017.png)

For 2018:

![VBA_Challenge_2018.png](/Resources/VBA_Challenge_2018.png)

Although we are able to gain more information from the refactored code, it takes the code longer to run. Twice as long in fact!

Here are the results for 2017 and 2018 respectively. It’s immediately apparent that 2017 was a more successful year than 2018 thanks to our formatting.

2017:

![VBA_Challenge_2017_Results.png](/Resources/VBA_Challenge_2017_Results.png)

2018:

![VBA_Challenge_2018_Results.png](/Resources/VBA_Challenge_2018_Results.png)


## In Summary

- **What are the advantages or disadvantages of refactoring code?**
* *Refactoring the code allows us to run one single script to analyze the yearly return of all stocks for 2017 and 2018. However, the code may seem much more confusing at first glance. And in terms of efficiency, the code does run twice as long, even if it is just by a quarter of a second.* *

- **How do these pros and cons apply to refactoring the original VBA script?**
* *If we are looking at much larger data sets, the efficiency of the code may decline with the refactored code.  It also took much more time to write and debug. By including the tickerIndex for the tickerVolume, there were many steps along the way that would allow for mistakes. However, once completed, the refactored code gives Steve all the information he needs to make safe decisions in purchasing stock from year to year!* *
