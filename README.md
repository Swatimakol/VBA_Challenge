# VBA Analysis Challenge

## Project Background
The purpose of this stock analysis is helping client obtain and compare *Total Daily Volume* and *Yearly Return* of each target stock in particular year.
By designing a VBA Macro for applying to different years, client can observe and make conclusion which stock performs better than others specifically in green energy industry.

We have designated 12 stocks to analyze in the year of 2018 and 2017. The worksheets are organized by year and contain the starting, final, highest, and lowest value of each stock for each day of the year. The VBA script will give the yearly increase/decrease in stock value for each stock, the percentage change over the year, and the total volume for the year. In addition the stock will give the greatest increase, decrease, and total volume.

The new refactored macro is created after keeping in mind to analyse huge data sets efficiently and quickly. It also gives an option take in differnet data sheets if need be, currenltly only 2018 and 2017 are available but in the future we can add more years and this macro will still be able to analyse for them. 

## Results

### Analysis comparision

In 2018, only *ENPH* and *RUN* two stocks had positive yearly Return as well as large Total Daily Volume. Both of them was outperformance than others green stocks.
You may also notice that *DQ* made it's return in negatives(-62%). While *AY* had the lowest daily return: 8,30,79,900

![](https://github.com/Swatimakol/VBA_Challenge/blob/3747a61ba51b887cbe12e7f114597f05c66b36f8/Resources/2018_results.png)

In 2017, all of stocks had positive Return except *TERP* (-7.2%). "DQ" made best yearly return with 199.4% but with lowest total Daily Volume (35,796,200) in 2017.

![](https://github.com/Swatimakol/VBA_Challenge/blob/3747a61ba51b887cbe12e7f114597f05c66b36f8/Resources/2017_results.png)

### Execution Times

For 2017 data set, it took 0.24seconds (rounded off) to complete the analysis

![](https://github.com/Swatimakol/VBA_Challenge/blob/3747a61ba51b887cbe12e7f114597f05c66b36f8/Resources/2017_timer.png)

For 2018 data set, it took 0.38seconds (rounded off) to complete the analysis

![](https://github.com/Swatimakol/VBA_Challenge/blob/3747a61ba51b887cbe12e7f114597f05c66b36f8/Resources/2018_timer.png)

## Summary

### What are the advantages or disadvantages of refactoring code?
   1. One of the advantage of refactoring the orignal code was it now can take huge data sets as oppose to taking the limited data.
   2. With the refactorred code we can now select which year we can want the data to run as oppose to orignal where it only took "2018".
   3. The refactored code places tickerIndex in an array which helpd sort out the runtime code faster while also using the memory in the contiguous manner. 
   4. The image below is the runtime of the orginal code placing "2018" worksheet - it took about 1.9 seconds while usign the refactor code it takes only 0.38 seconds. Speeding up the analysis with more than 50%.




![](https://github.com/Swatimakol/VBA_Challenge/blob/f969b250a2c450e2a4d52da82a2bdcd1d5bef6ca/Resources/OrignalCode_timer.png)


  5. One big *Disadvanatge* I see with this script is that it can only take 12 defined stocks. If the use decide to add another stock, this script witll not work for it. In that case we need to make alterations.


### How do these pros and cons apply to refactoring the original VBA script?

   *We followed a flow that helped achieve efficient code:*
    
   1. First we asked the user to place in the year we want to work in - here they can place in 2017 or 2018, if not one of these we will have a message indicatiing to place only between 2018 or 2017.
   
  ```
    If yearValue = "2018" Or yearValue = "2017" Then
    
         Worksheets("All Stocks Analysis").Activate
    
    Else
         MsgBox ("Kindly input correct year value as 2017 or 2018")
    
    End If
  ```
  
  2. A nested loop was created, that went throughstocks original data and retrieve ticker name, startingPrices and endingPrices, and save information to        each related tickerIndex.
  
   ```
     For i = 2 To RowCount
     
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If

         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            tickerIndex = tickerIndex + 1
         End If
    
    Next i
   ```
   
   3. As you observe in the above code, while we are getting the starting and the ending price of te ticker, we also storing the total value of every ticker under the same nested loop using tickerVolumes as the variable. 
   
   ```
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
   ```
    
   4. We also used loops to place analysis of Ticker name, %return and Total volume.
   
   ```
     For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
     Next i
    
   ```
   
   5. To make the code look presentable, we formatting within in the same macro instead of using a different macro. Helping reduce the runtime. Here we also made use of the efficient looping to assign colors to the cells if the return value is below 0 or above. 
  
   6. Lastly I added Two buttons - Enetr the year and Clear data - so we can communicate/run the data while being on the excel file using GUI. 
   

     
    
  
