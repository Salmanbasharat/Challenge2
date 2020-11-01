# ***Challenge2- Stock Prices Analysis***
## **Overview of Project**
The purpose of this project was to get the true picture of the performance of certain stocks (represented by their Ticker symbols) in the stock market for the years 2017 and 2018 and analysis was done on 12 stocks to see their **Total Volumes** and **Returns** for the years under consideration. 
## **Results of the Analysis**
The Total Volumes and Returns of all the 12 stocks were calculated for both the year 2017 and 2018. The VB codes used to calculate the *Total Volumes* and *Returns* are given below:



'''
'1a) Create a ticker Index
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
        
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For j = 0 To 11
        tickerVolumes(j) = 0
        Next j
        
     ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
            
        '3a) Increase volume for current ticker
        
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                    
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                     
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                             
           
        'End If
        
            End If
            
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
            
            '3d Increase the tickerIndex.
            
                         
                tickerIndex = tickerIndex + 1
                
        'End If
            
            End If
       
                
    Next i
    
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
                
    Next i
    '''




The result of these analysis using the VB codes are given in the below pictures.
### **2017**

<img src="2017 Analysis_Table.PNG"><img>
### **2018**
<img src="2018 Analysis_Table.PNG"><img>

The *Daily Volumes* of all of these stocks in 2017 was about 3.1 million whereas in 2018 it was 3.3 million which is not a big difference however the volumes of individual stocks have significantly changed for example if we take the example of **ENPH** the trading in this stock is almost 3 times of 2017 volumes.

This is visble that all these stocks performed exceptionally well in 2017 as compared to 2018. In 2017 only one stock **TERP** was in negative territory highlighted by red color whereas all the other stocks were in green color meaning giving positive returns. These returns have been calculated by dividing the year ending prices of these stocks with year starting prices.  
In 2018 except **ENPH** and **RUN** all the stocks have given negative returns however it is interesting to note that the return of even **ENPH** from 2017 levels has decreased from 129.5% to 81.9% whereas the return of **RUN** is not only in positive territory but also increased significantly in 2018 to 84% from 5.5% of 2017. The trading volume of this stock has increased by almost 2.5 times which shows the interest in this stock by the investors. So by looking at this data, **RUN** is a good stock to invest in.

## **Result of Refactoring**

The analysis was performed in VBA with and without **Refactoring**.  There is a significant drop of execution time once *Refactoring* was done. The execution times in both cases are shown in below pictures.
### **Original Script**

<img src="2017 run time wo refactoring.PNG"><img>
<img src="2018 run time wo refactoring.PNG"><img>

### **Refactored Script**
<img src="2017 Refactored execution time.PNG"><img>
<img src="2018 Refactored execution time.PNG"><img>

This is visible from above pictures that execution time in 2017 analysis has reduced from 0.97 seconds to 0.15 seconds. In the year 2018 it has reduced from 0.94 sec to 0.17 sec. So refcatoring has helped in this case.
## **Summary**
I have used *Refactoring* in simplifying the code. In general there are many advantages and some disadvantages of using Refactoring .
### **Advantages**
In the words of Martin Fowler (Father of Code Smell), below are the advantages:

*Refactoring Improves the Design of Software

*Refactoring Makes Software Easier to Understand

*Refactoring Helps Finding Bugs

*Refactoring Helps Programming Faster

However there could be some disadvantages as well of using Refactoring.
### **Disadvantages**
*It is time consuming

*It may cause more money

*It is risky when application is big

### **Advanatges & Disadvantages in this Analsyis**
The major difference or Advantage of using the Refactored code in this case is that now we are able to expand our analysis as compared to the original script, Even we can expand the analysis beyond 2017 and 2018.

The other advatage is that Refactored code has taken significantly lesser time and have used much lesser resources as compared to the original script because it has looped over the whole data in one time. 


