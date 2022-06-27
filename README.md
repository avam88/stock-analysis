# stock-analysis

# Scalable VBA Script for 2017 & 2018 Stocks Analysis
## 2 analysis. original analysis from our client to deliver an interactive vba script that would allow analysis of all stocks of a selected year with a click of a button. Our original code worked great, but our client expanded the scope of their ask. The client is now curious about . We then needed to revisit our code to make sure that it was scalable. Could our current vba script that worked on 12 stocks for 2 years of analysis be effectively used for all stocks over many years? The 2nd deliverable in this project is an refactoring of our original code. 

# VBA Code for original analysis
## The first task in this project was to script code to deliver outputs on total daily volume and annual return for stocks in either 2017 or 2018 financial years. 
### our primary tool to deliver run this analysis was using nested loops. We wanted to first run through each. . in essence the logic flow of our code in plain text

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
### the workhorse in our code is the nexted for loop and conditional statements. (insert photo of code). 

### The above language asks our analysis mechanism to first run through all of our ticker names (stock names), and for each ticker loop through each row in our data set to find the values that meet the crieteria. dive deeper and sum the values of all trading. Then we want to find the closing price at the beginning of the year and the closing price at the end of the year. The percent difference between these values delivers our annual rate of return. WithiThis is a powerful analysis that allows our client to make informed financial decisions for their clients in turn.

# VBA Code Refactored
## After delivery of our first project analysis our client asked if they could use this same mechanism to analyse all stocks for any year. While our script ran effectively 
### In order to grab a baseline funcationality we added a timer function to our code. Now anytime our macro was initiated it would deliver 2 outputs; 1) the original stock analysis and 2)a timestampe of how long it took the computer to run the analysis. In order to do this we defined our timer variables and created a function that ran outside of our stock analysis loop. Our original code ran the analysis for each year in just under a second.

insert original photos here.

### While functionality for our original purposes of limited set of stocks for 2 year was high, you can image if we increase the size of our data set by a magnitude of 100 or 1000, this code could potentially take a very long time to run. If we zoom out, We set about refactoring or editing our code to see if we could find more better structure/architecture for a more efficient loop. This is what we came up with.

### here are the new run times. 
insert new timestamp photos.
### we were able to cut down our run times by almost 3/4. 

# Summary
- original code is user friendly. The script is much more readable and digestible to non code writers
- refactored code runs more efficiently
- 
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).


