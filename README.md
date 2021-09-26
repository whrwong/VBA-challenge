# Vba Challenge
UCI BC VBA scripting assignment 9/25/21

This code loops through three sheets of stock data, one sheet per year for the years 2014, 2015, and 2016.
The input data includes ticker symbols, dates, opening prices, high prices, low prices, closing prices, and the volume (shares).
The code I wrote creates four columns corresponding to: unique ticker symbol, yearly change and percent change (from the latest closing price to the earliest opening price), and cummulative volume per ticker.

* Note that each volume was divided by 1000 to allow VBA to run on the data.
The long integer that holds the "volume" variable can only hold up to 2,147,483,647 but the second ticker AA has a volume of 5,205,551,000 for 2016 make it impossible for VBA to handle the data.
