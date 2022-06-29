# A Scalable & Interractive VBA Macro For Stock Market Analysis
### Using VBA macros in the application of Excel, we will deliver an analysis mechanism that will help financial advisors efficiently return information about stocks in order to make executive investing decisions. The original brief of this project was to deliver an interactive vba script that would allow analysis of a small set of stocks for a selected year with a click of a button. The macro analysis needed to compare two major elements of stock performance across a small subset of stocks; the total daily volume (the number of times a stock is traded in a day which is a reflection of activity and general health) and the annual rate of return. After designing the analysis mechanism using vba code, the second phase of the project was to increase the demand load on our code - could we use the script as it was currently written to perform analysis on all stocks for any year? Our task then was to revisit and refactor our code and assess whether it was scalable - could our current vba script that worked on 12 stocks for 2 years of analysis be effectively used for all stocks over many years? 

## VBA Code Stock Analysis for 2017 & 2018
### The first task in this project was to write a VBA macro to deliver outputs on total daily volume and annual rate of return for 12 separate stocks in either 2017 or 2018 financial years. 
In order to write coherent code we wanted to outline our logic flow to ensure our order of operations was correct. 
- Declare an interractive variable for user input to define analysis year.
- Create and format an output sheet titled "All Stocks Analysis".
- Initialize an array of all the stock names - 'tickers'.
- Initialize variables for the starting price and ending price.
- Activate the worksheet we want to access data from (2017 or 2018 depending on user input).
- Find the number of rows to loop over.
- Loop through the 'ticker' array.
- Within this loop, loop through all the rows in the data.
- Find the total volume for the current ticker.
- Find the starting price for the current ticker.
- Find the ending price for the current ticker.
- Output the data for the current ticker.

our primary tools to run this analysis were "nested for loops" and conditional statements.

![nested for loops original](https://user-images.githubusercontent.com/107326987/176092315-c7a8180a-8d40-4c99-9e37-28abbe9bc891.png)

The above language asks our analysis mechanism to first run through the variable that we declared to hold the stock names - all the names of the stocks in our dataset are assigned as values to the 'ticker' variable. The opening line of code in the above image is the first 'for loop' - we are asking the macro to run through every stock name in the 2017/2018 worksheets. Then we enter our 'nested for loop' which asks our macro to loop through each row in our data set to find everywhere that our specific 'ticker' value appears. When the macro finds a row with that specific 'ticker' value, then we enter into our conditional statements (highlighted in orange). Our conditional statements use the "if/then" function to check first if a certain criteria is met to then grab information from corresponding cells to output at the end of the loop. Our conditional statements allow us to find the closing price at the beginning of the year and the closing price at the end of the year and the total daily volume for each 'ticker' and ultimately let us perform a function to find the annual rate of return. This is a powerful analysis that allows our client to make informed financial decisions for their clients in turn.

## VBA Code Refactored
### After adding an interactive button to our excel file to allow for input (the year we want run analysis for), our macro returns the analysis successfully. Our client asked, however, if this same mechanism could be used to analyze all stocks for any year. While our script ran effectively for a smaller dataset, we need to assess the architecture of our code to determine if it can efficiently analyze much larger data files.
In order to grab a baseline funcationality we added a timer function to our code. Now any time our macro was initiated it would deliver 2 outputs; 1) the original stock analysis and 2)a timestamp of how long it took the computer to run the analysis. In order to do this we defined our timer variables and created a function that ran outside of our stock analysis loop. Our original code ran the analysis for each year in just under a second.

### Original Code Run Time
![VBA_Challenge_2017_original](https://user-images.githubusercontent.com/107326987/175865290-bd430456-4c34-46f9-93e0-c16f8eafb783.png)
![VBA_Challenge_2018_original](https://user-images.githubusercontent.com/107326987/175865301-99c3913a-9363-4e84-83d8-cd81bc4c1f04.png)

While functionality for our original purposes (limited set of stocks, limited number of years) is adequate, we can extrapolate that if we increase the size of our data set by a magnitude of 100 or 1000, this code could potentially take an exorbitant amount of time to run. We set about refactoring or editing our code to see if we could find better structure, syntax and logic flow for a more efficient loop through the data.
Ultimately it is cumbersome to ask the macro to loop through everyline of the dataset (currently all 3013 rows with values) for each stock name ('ticker'). With 12 separate stocks, we're essentially running through 36,000+ rows of data. What if instead, we could run through our rows just once and grab all the values based on our conditional statements to return the information we need?

![refactored for loop](https://user-images.githubusercontent.com/107326987/176099015-d1a3f53b-8127-487c-a9c2-66dffaa62d7b.png)

The original code had our conditional statement start immediately inside our nested loop, essentially asking "does this current row have the ticker value (i)?". If the statement is false, then the loop moves to the next row. The macro works its way through every row, and then loops back to the begining row for the next ticker value. We can see, however, that our stock data is organized by 'ticker' name, and within that, chronologically. If we declare the variable tickerIndex = 0 and move through the rows sequentially, we are moving through ticker "AY" and summate the values in the "Total Daily Volume" column. Then we enter into our conditional statements that set the parameters to allow us to find the end of the current ticker subset and the beginning of the next ticker value (when the row above/below no longer contains the same value as the current row). If that condition is met we either grab the starting/ending price value. Lastly, if that condition is met for ending price we increase our ticker index by 1 and move to the next subset of stocks. Our refactored code only touches each row once, not wasting unecessary time touching every row for every stock name. If we declare our outputs as arrays outside of our loops the code will earmark the values to a ticker and output all the those values to predefined cells corresponding to the ticker value.

### Refactored Code Run Time 
![VBA_Challenge_2017](https://user-images.githubusercontent.com/107326987/175865335-6d07687a-772c-4a72-8c2c-09fff7ebde7e.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/107326987/175865349-6310eff6-86bd-468b-b565-efa6f76962eb.png)
Through refactoring we were able to cut down our run times by roughly 85%. 

## Summary
- original code is user friendly. The script is much more readable and digestible to non code writers and it works!
- refactored code runs more efficiently. however refactoring is an opportunity to pollute working code with bugs. And ultimately, we've made our code more efficient, but excel, vba and macros are not optimized to perform analysis on large sets of data. Ultimately our refactored code will still be less efficient than if we wrote this in a different language different platform.

Specifically our refactored code is worse/better because . . .
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt). Both rely on the fact that 


