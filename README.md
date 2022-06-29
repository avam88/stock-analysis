# A Scalable & Interactive VBA Macro For Stock Market Analysis
### Using VBA macros in the application of Excel, we will deliver an analysis mechanism that will help financial advisors efficiently return information about stocks in order to make executive investing decisions. The original brief of this project was to deliver an interactive VBA script that would allow analysis of a small set of stocks for a selected year with a click of a button. The macro analysis needed to compare two major elements of stock performance across a small subset of stocks; the total daily volume (the number of times a stock is traded in a day which is a reflection of activity and general health) and the annual rate of return. After designing the analysis mechanism using VBA code, the second phase of the project was to increase the demand load on our code - could we use the script as it was currently written to perform analysis on all stocks for any year? Our task then was to revisit and refactor our code and assess whether it was scalable to effectively analyze a much larger dataset. 

## VBA Macro : Stock Analysis for 2017 & 2018
### The first task in this project was to write a VBA macro to deliver outputs on total daily volume and annual rate of return for 12 separate stocks in either 2017 or 2018 financial years. 
Our primary tools to run this analysis were "nested for loops" and conditional statements.

![nested for loops original](https://user-images.githubusercontent.com/107326987/176092315-c7a8180a-8d40-4c99-9e37-28abbe9bc891.png)

The above language asks our analysis mechanism to first run through the variable that we declared to hold the stock names - all the names of the stocks in our dataset are assigned as values to the 'ticker' variable. Before beginning our for loops, we declare an interractive variable that allows the macro user to define which year to run analysis for. The opening line of code in the above image is the first 'for loop' - we are asking the macro to run through every stock name in either the 2017 or 2018 worksheet. Then we enter our 'nested for loop' which asks our macro to loop through each row in the worksheet to find everywhere that our specific 'ticker' value appears. When the macro finds a row with that specific 'ticker' value, then we enter into our conditional statements (highlighted in orange). Our conditional statements use the "if/then" function to check first if a certain criteria is met to then grab information from corresponding cells to output at the end of the loop. Our conditional statements allow us to find the closing price at the beginning of the year and the closing price at the end of the year and the total daily volume for each 'ticker' and ultimately let us perform a mathematical function on our defined variables to find the annual rate of return. Ultimately the code contains language to format our output cells depending on their values for immediate readability. This is a powerful analysis that allows our client to make informed financial decisions for their clients in turn.

Using our VBA script to run the analysis, our client can see that these stocks performances vary widely from year to year. 

![Screen Shot 2022-06-28 at 8 11 33 PM](https://user-images.githubusercontent.com/107326987/176343210-e01e580f-7e8b-4a8f-bef4-8f1ee3bba7a4.png)
![Screen Shot 2022-06-28 at 8 11 47 PM](https://user-images.githubusercontent.com/107326987/176343218-65492aa4-68cf-45b0-ac79-9340d6a8614a.png)


## VBA Code Refactored
### While our script ran effectively for a smaller dataset, we need to assess the architecture of our code to determine if it can efficiently analyze much larger data files. Our brief now is to determine whether we can restructure our code to perform the same analysis more quickly and/or using lessing memory.
In order to grab our baseline functionality we added a timer function to our code. Now any time our macro is initiated it will deliver 2 outputs; 1) the original stock analysis and 2) a timestamp of how long it took the computer to run the analysis. In order to do this we defined our timer variables and created a function that ran outside of our stock analysis loop. Our original code ran the analysis for each year in just under a second.

### Original Code Run Time
![VBA_Challenge_2017_original](https://user-images.githubusercontent.com/107326987/175865290-bd430456-4c34-46f9-93e0-c16f8eafb783.png)
![VBA_Challenge_2018_original](https://user-images.githubusercontent.com/107326987/175865301-99c3913a-9363-4e84-83d8-cd81bc4c1f04.png)

While functionality for our original purposes (limited set of stocks, limited number of years) is adequate, we can extrapolate that if we increase the size of our data set by a magnitude of 100 or 1000, this code could potentially take an exorbitant amount of time and memory to run. We set about refactoring our code to see if we could find better structure, syntax and logic flow for a more efficient loop through the data.
Our original macro loops through every line of the dataset (currently all 3013 rows with values) for each stock name ('ticker') which is cumbersome. With 12 separate stocks, we're essentially running through 36,000+ rows of data. What if instead, we could run through our rows just once and grab all the values based on conditional statements to return the information we need?

![refactored for loop](https://user-images.githubusercontent.com/107326987/176339694-6914d28f-bd6a-4786-9414-5bef9aef8049.png)

The original code had our conditional statement start immediately inside our nested loop, essentially asking "does this current row have the ticker value (i)?". If the statement is false, then the loop moves to the next row. The macro works its way through every row, and then loops back to the beginning row for the next ticker value. We can see, however, that our stock data is organized by 'ticker' name, and within that, chronologically. If we declare the variable tickerIndex = 0 and move through the rows sequentially, we begin by moving through ticker "AY" and summate the values in the "Total Daily Volume" column. Then we enter into our conditional statements (highlighted in pink) that set the parameters to allow us to find the end of the current ticker subset and the beginning of the next ticker value (when the row above/below no longer contains the same value as the current row). If that condition is met we either grab the starting/ending price value respectively. Lastly, if that condition is met for the ending price conditional statement, we increase our tickerIndex by 1 and move to the next subset of stocks as seen in the code line highlighted in orange.

![Screen Shot 2022-06-28 at 7 37 35 PM](https://user-images.githubusercontent.com/107326987/176341497-c11a6b50-7f1c-4886-b781-ab48556c90e4.png)

![Screen Shot 2022-06-28 at 7 37 51 PM](https://user-images.githubusercontent.com/107326987/176341507-5c08acdf-bbd5-43be-9cd2-353c09affe13.png)

Our refactored code only touches each row once, not wasting unnecessary time touching every row for every stock name. If we declare our outputs as arrays outside of our loops the code will earmark the values for our defined ticker and output all the those values to predefined cells corresponding to the ticker value. The first image shows our output variables declared as arrays before our for loop. The second image illustrates the formatting of our array variable outputs. Through refactoring we were able to cut down our run times by roughly 85%. 


### Refactored Code Run Time
![VBA_Challenge_2017](https://user-images.githubusercontent.com/107326987/175865335-6d07687a-772c-4a72-8c2c-09fff7ebde7e.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/107326987/175865349-6310eff6-86bd-468b-b565-efa6f76962eb.png)

## Summary : Why Refactor?
Our original code worked, so why refactor? We can see here that with careful re-analysis of our dataset we can identify patterns that allow us to restructure our logic flow or syntax to deliver code that is more efficient. There are, however, drawbacks to refactoring code. We can borrow from the sentiment "if it isn't broke, don't fix it" - refactoring code creates an opportunity to pollute working code with bugs and functions that error out. In this case specifically, our refactored code is more efficient, but our original code is much more readable and digestible to non code writers. 

Ultimately, both of our codes work exclusively because of the way the dataset is organized. If our data were to be jumbled (no longer organized first by stock name and in turn chronologically) our macro code would not return accurate values. Excel is also not the most powerful tool to perform analysis on large datasets. While our refactored code is more efficient, excel is not the optimal application to work on a large dataset like one that would contain information for all stocks and all years.  Ultimately our refactored VBA code will still be less efficient than written in a different language using different data analysis application.



