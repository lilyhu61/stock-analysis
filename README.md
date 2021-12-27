# ALL STOCKS ANALYSIS REFACTORED CHALLENGE WITH VBA

## OVERVIEW OF PROJECT
### Purpose

In this project, we try to write VBA solution code for Steve that click of a button, he can analyze an entire dataset, and he can expand the dataset to conclude the entire stock market over the last few years. Although the code should work well for a dozen stocks and for thousands of stocks as well, and should reduce execution time. In this challenge, we will edit, or refactor, make the code more efficient Ð by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.

### Our challenge Data Background
Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.

Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.
## RESULTS: Refactor VBA code and Measure Performance
### Deliverable Requirements, Code Examples, Compare Stock Performance and Timestamp procedure below:
**1. The tickerIndex is set equal to zero before looping over the rows.**
 ```
     '1a) Create a ticker Index
     
       For i = 0 To 11
          tickerIndex = tickers(i)
```  
**2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.**
 ```
      '1b) Create three output arrays
       Dim tickerVolumes As Long
       Dim tickerStartingPrice As Single
       Dim tickerEndingPrice As Single
 ```   
**3. The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.**
 ``` 
    Worksheets(yearValue).Activate
    tickerVolumes = 0
        
        '2b) Loop over all the rows in the spreadsheet.
        For j = 2 To RowCount
    
           '3a) Increase volume for current ticker
            
              If Cells(j, 1) = tickerIndex Then
  ``` 
**4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.**
 
**5. Code for formatting the cells in the spreadsheet is working.**

**6. There are comments to explain the purpose of the code.**

**7. The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module.**

   >***Dataset Examples provided***
   
   ![image](https://user-images.githubusercontent.com/95242493/147422564-cafccfcc-e811-472e-8ef6-cdbada852540.png)
   ![image](https://user-images.githubusercontent.com/95242493/147422727-8941b3d3-cc6c-432e-ac3c-6f77004d5b25.png) 
   
   
   >***Final VBA Analysis 2017***
   
    All Stocks (2017)		
		
![image](https://user-images.githubusercontent.com/95242493/147435578-9fbb4e83-9fe7-4b70-ad34-c7018052b513.png)


   >***Final VBA Analysis 2018***
		
     All Stocks (2018)		
![image](https://user-images.githubusercontent.com/95242493/147435681-8e4b8680-9263-43e6-826d-bc7cb7ee6c39.png)

**8. The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2017.png and VBA_Challenge_2018.png**
   > Running our fully 2017 and 2018 data stock analysis gave us an elapsed run time for each year, below our results.
   
   > ***Time on VBA_Challenge_2017.PNG***
   
   ![Screen Shot 2021-12-26 at 9 58 24 PM](https://user-images.githubusercontent.com/95242493/147436061-a6b2c481-8e34-4485-a7ea-5448e9ca5859.png) 
   
   > ***Time on VBA_Challenge_2018.PNG***


   ![Screen Shot 2021-12-26 at 10 04 42 PM](https://user-images.githubusercontent.com/95242493/147436353-e927cb2b-ad22-43fb-a0a4-822bc4479a40.png)

   
  
    
 ## SUMMARY
 ### 1. What are the advantages or disadvantages of refactoring code? 
 ###    Advantages:
     - Make the VBA script run faster.
     - Make the code more efficient, using less memory.
     - Improving the logic of the code to make it easier for future users to read.  
     - Maintainability and scalability.    
###    Disadvantages:
    - Might have to retest lots of functionality.
    - Refactoring process can affect the testing outcomes.
### 2. How do these pros and cons apply to refactoring the original VBA script? 
       Code Refactoring is an important exercise to remove code smell. It helps to find bugs, makes programs run 
       faster, it's easier to understand the code, improves the design of software, etc. Code smell slows down 
       the development and is prone to more defects. An adequate set of unit tests and a supportive environment
       should be there for code refactoring.
