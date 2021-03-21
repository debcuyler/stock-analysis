# Stock-analysis for Steve

## Overview of Project

### Purpose
The purpose of this project was to help Steve analyze stock data total volumes and returns from twelve different companies in the years 2017 and 2018. Steve is helping his parents decide which tech stock would be a good investment. His parents are choosing based on name, DQ, because they like the restaurant with the same name. Steve would rather analyze the volumes traded and returns on a number of tech stocks to ensure his parents are making a good investment choice.

## Results
 - Results of stock analysis: When comparing the volumes traded and return rate on twelve different tech companies, we can see that eleven of the twelve companies had a positive growth return in 2017. Of those same companies, only two of them had a positive growth return in 2018. Although "DQ" had a 199% return rate in 2017, in 2018 they had a 62% loss. The trade volumes also increased three times in 2018 as compared to 2017. When analyzing the data from all twelve stocks, two of them had positive return rates in both 2017 and 2018. The ticker "RUN' had double their daily volume in 2018 as they had in 2017 and their return rate increased from 5.5% in 2017 to 84% in 2018. In comparison, ticker "ENPH" almost tripled their daily volume from 2017 to 2018, but their rate of return, although positive, dropped from 130% in 2017 to 82% in 2018. 
 - The code was refactored from it's original state by using output arrays and referencing the tickerIndex in order to run the code more efficiently. If I had simply re-used the original code that was done only for "DQ", then each ticker would have held it's own loop of code and the code would have been copied over many times. Using the index allowed me to write the code once and loop through all of the tickers as well as all of the rows of data in the spreadsheet.

## Summary

- What are the advantages or disadvantages of refactoring code? The advantages of refactoring code is that the code is more efficient and runs faster. It also makes the code more understandable and assists in finding bugs. The disadvantages of refactoring code are that it can be very time consuming and it can also take away from development time for other projects. Also, even if a lot of time is spent refactoring the code, it may produce a non-optimal result.    

- How do these pros and cons apply to refactoring the original VBA script? The codes runs much faster for sure, twice as fast to be specific. When running code on a large amount of data, time matters. The con is that the code worked perfectly fine without refactoring and it took a long time to re-write the code. I also ran into a lot of issues with errors and syntax and could not get the code to run at all for a long time. 