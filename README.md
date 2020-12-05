# Stock Analysis with VBA

## Overview of Project

### Purpose

The purpose of this analysis is to help Steven compare twelve stocks over the years 2017 and 2018. The key information is the daily volume of the stocks and their returns. His parents were iniitally only interested in the stock "DQ," so showing all the stocks together can determine if that is the best stock for them.

### Code

Using VBA in Excel, a macro was used to collect data of stocks based on year. The macro will first ask the user which year to run on using an `InputBox`; the yearly results are seperated by Worksheet. All the stocks are already sorted in alphatebetical order. Once the year is decided, the macros will go through each row, and determine what stock it is looking through by seeing if the stock is the same as the one before it. This conditional nested for-loop is shown below:

```
    For i = tickerIndex To 11
        
        ticker = tickers(i)
        tickerVolumes(i) = 0
        'MsgBox i
    
        
        ''2b) Loop over all the rows in the spreadsheet.
        For j = 2 To RowCount
        
            If Cells(j, 1).Value = ticker Then

               '3a) Increase volume for current ticker
               tickerVolumes(i) = tickerVolumes(i) + Cells(j, 8).Value

           End If
      
            '3b) Check if the current row is the first row with the selected tickerIndex.
            'If  Then
            If Cells(j, 1).Value = tickers(i) And Cells(j - 1, 1).Value <> tickers(i) Then
                
                tickerStartingPrices(i) = Cells(j, 6).Value
                
            'End If
            End If
            
            '3c) check if the current row is the last row with the selected ticker
             'If the next row’s ticker doesn’t match, increase the tickerIndex.
            'If  Then
            If Cells(j + 1, 1).Value <> tickers(i) And Cells(j, 1).Value = tickers(i) Then
    
                '3d Increase the tickerIndex.
                tickerEndingPrices(i) = Cells(j, 6).Value
        
           
                
            'End If
            End If
        
        Next j
    
    Next i

```
This macro was refractored to store all information in arrays. This allows itself to only need to loop through the rows once instead of three different times. 

Once the macro runs through all each row, it will input the values into another sheet. This is done by looping through the new arrays. The formula to determine the Return value used is `tickerEndingPrices(i) / tickerStartingPrices(i) - 1`. Conditional formatting is then done for positive and negative returns. After refractoring the macro, both 2017 and 2018 took just under half a second to run, as shown by their end `MsgBox` values:

![2017 MsgBox](https://github.com/ajg318/stock-analysis/blob/main/resources/2017runtime.png)

![2018 MsgBox](https://github.com/ajg318/stock-analysis/blob/main/resources/2018runtime.png)

This is about 20% faster than the original script with multiple loops for each value type.

## Analysis and Challenges

### Analysis of Outcomes Based on Year

A focus for Steven was to determine if DQ is the best stock for his parents. In 2017, this stock boasted the largest return with 199.4% gain. It was also the smallest stock in daily volume. In 2018, DQ had the highest percent loss in returns while growing in daily volume. Being such a volatile stock, Steven's parents will need to determine if the risk is worth the reward.

In general, 2017 was a much more successful year for these stocks than 2018. In 2017, only one stock finished with a negative return while in 2018 every stock was negative besides two. ENPH and RUN were positive each year and TERP was negative. Using these two years as evidence, ENPH and RUN seem to be the safest of the lucrative stocks while DQ and SEDG are the most volatile. 

Here are the results to 2017 and 2018, respectively:

![2017 Results](https://github.com/ajg318/stock-analysis/blob/main/resources/2017values.png)

![2018 Results](https://github.com/ajg318/stock-analysis/blob/main/resources/2018values.png)

### Challenges and Difficulties Encountered

The dataset used was clean before needing any attention. This limited the challenges faced when using it. One challenge I did face was using the correct values and types for the arrays and variables. I initially was having overflow errors. Once I chose the correct values, however, this was resolved. 

## Results

Refractoring code has advantages and disadvantages. Some advantages are:

*It makes the code more efficient. This macro collected the same data in a faster amount of time (about 20% faster)
*It makes the code more reusable. Adding comments and removing magic numbers allows others (and you) to be able to go back to this code and incorporate more to it. This macro has some functions I have used to work on the assessment and work globally now.
*It future proofs the code. If the data grows or more years are needed, the code can now be available to more than this situation. This macro went from only looking from DQ to looking at twelve stocks.

Some disadvantages are:

*It will take time to refractor. If a project has a priority of budget and hours over efficiency, going back to work on the code may not be possible. This macro did not take too much time to refractor nor am I on a budget, so this was not a problem.
*It can possibly break the functionality of the code. This should not happen but trying to fix something that is not broken can cause damage. Ultimately, however, the code should end up with the same use. Initially the macro had bugs, but these were resolved.



