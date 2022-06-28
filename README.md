# Scalable & Interractive VBA Macro For Stock Market Analysis
### Using VBA language and macros in the application of Excel, we will deliver an analysis mechanism that will help financial advisors efficiently return information about stocks in order to make executive investing decisions. The original brief of this project was to deliver an interactive vba script that would allow analysis of a small set of stocks for a selected year with a click of a button. The macro analysis needed to compare two major elements of stock performance across a small subset of stocks; the total daily volume (the number of times a strock is traded in a day which is a reflection of activity and general health) and the annual rate of return. After designing the analysis mechanism using vba code, the second phase of the project was to increase the demand load on our code - could we use the script as it was currently written to perform analysis on all stocks for any year? Our task then was to revisit our code and assess whether it was scalable - could our current vba script that worked on 12 stocks for 2 years of analysis be effectively used for all stocks over many years? 

## VBA Code Stock Analysis for 2017 & 2018
### The first task in this project was to script code to deliver outputs on total daily volume and annual rate of return for stocks in either 2017 or 2018 financial years. 
In order to write cohesive code we wanted to outline our logic flow to ensure our order of operations was correct. 
- Format the output sheet on the "All Stocks Analysis" worksheet.
- Initialize an array of all tickers.
- Prepare for the analysis of tickers.
- Initialize variables for the starting price and ending price.
- Activate the data worksheet.
- Find the number of rows to loop over.
- Loop through the tickers.
- Loop through rows in the data.
- Find the total volume for the current ticker.
- Find the starting price for the current ticker.
- Find the ending price for the current ticker.
- Output the data for the current ticker.
our primary tools to run this analysis were "nested for loops" and conditional statements - the workhorses of this mechanism.

![nested for loops original](https://user-images.githubusercontent.com/107326987/176092315-c7a8180a-8d40-4c99-9e37-28abbe9bc891.png)

The above language asks our analysis mechanism to first run through the variable that we declared to hold the stock names - all the names of the stocks in our dataset are assigned as values to the 'ticker' variable. The opening line of code in the above image is the first 'for loop' - we are asking the macro to run through every stock name in the 2017/2018 worksheets. Then we enter our 'nested for loop' which asks our macro to loop through each row in our data set to find everywhere that our specific 'ticker' value appears. When the macro finds a row with that specific 'ticker' value, then we enter into our conditional statements (highlighted in orange). Our conditional statements use the "if/then" function to check first if a certain criteria is met to then grab information from corresponding cells to output at the end of the loop. Our conditional statements allow us to find the closing price at the beginning of the year and the closing price at the end of the year and the total daily volume for each 'ticker' and ultimately let us perform a function to find the annual rate of return. This is a powerful analysis that allows our client to make informed financial decisions for their clients in turn.

## VBA Code Refactored
### After adding an interactive button to our excel file to allow for input (the year we want run analysis for), our macro returns the analysis successfully. Our client asked, however, if this same mechanism could be used to analyse all stocks for any year. While our script ran effectively for a smaller dataset, we need to assess the architecture of our code to determine if it can efficiently analyze much larger data files.
In order to grab a baseline funcationality we added a timer function to our code. Now any time our macro was initiated it would deliver 2 outputs; 1) the original stock analysis and 2)a timestampe of how long it took the computer to run the analysis. In order to do this we defined our timer variables and created a function that ran outside of our stock analysis loop. Our original code ran the analysis for each year in just under a second.

### Original Code Run Time
![VBA_Challenge_2017_original](https://user-images.githubusercontent.com/107326987/175865290-bd430456-4c34-46f9-93e0-c16f8eafb783.png)
![VBA_Challenge_2018_original](https://user-images.githubusercontent.com/107326987/175865301-99c3913a-9363-4e84-83d8-cd81bc4c1f04.png)

While functionality for our original purposes (limited set of stocks, limited number of years) is adequate, we can extrapolate that if we increase the size of our data set by a magnitude of 100 or 1000, this code could potentially take an exorbitant amount of time to run. We set about refactoring or editing our code to see if we could find better structure, syntax and logic flow for a more efficient loop through the data.
Ultimately it is cumbersome to ask the macro to loop through everyline of the dataset (currently all 3013 rows with values) for each stock name ('ticker'). With 12 separate stocks, we're essentially running through 36,000+ rows of data. What if instead, we could run through our rows just once and grab all the values based on our conditional statements to return the information we need?

INSERT IMAGE OF NEW CODE HERE

instead of running through the entire code for every ticker to obtain the output information for 1 ticker before starting over again. We are asking the code to run through each line and identify the ticker, then initiate the for loop to run the conditional statements. This way the code only touches each line once, not wasting unecessary time touching each line for every value of stoc, name. IS THIS TRUE???? now we are claiming our outputs as arrays outside of our loops. So the code now tabulates for each row, earmarking the values to a ticker and then once the loop through the rows is finished we loop through our ticker values again to output all the new values to cells corresponding to ticker value.

### Refactored Code Run Time 
![VBA_Challenge_2017](https://user-images.githubusercontent.com/107326987/175865335-6d07687a-772c-4a72-8c2c-09fff7ebde7e.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/107326987/175865349-6310eff6-86bd-468b-b565-efa6f76962eb.png)
Through refactoring we were able to cut down our run times by roughly 85%. 

## Summary
- original code is user friendly. The script is much more readable and digestible to non code writers and it works!
- refactored code runs more efficiently. however refactoring is an opportunity to pollute working code with bugs. And ultimately, we've made our code more efficient, but excel, vba and macros are not optimized to perform analysis on large sets of data. Ultimately our refactored code will still be less efficient than if we wrote this in a different language different platform.

Specifically our refactored code is worse/better because . . .
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).


