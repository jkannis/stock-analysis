# stock-analysis

## Overview of Project
The client for this project was asked to assist in analyzing stock performance for investment purposes. This project refactors code used to calculate the total daily volume and annual return on stocks for a selected year to present meaningful information for decision-making for the client. This analysis was performed on the provided dataset using Excel and VBA, and refactoring provided code for improved performance. The daily volume was accumulated for each stock ticker and the return was calculated using the the starting and ending prices for each ticker. The results were stored in an array during processing and written to the output sheet after all processing was complete.
### Results
- The original script used to process the stocks data used logic that looped through the dataset looking for a specified ticker, then calculated the volume and return for that ticker and wrote it to the sheet. Then the logic was repeated for each ticker in the array.
  - This process was successful in providing accurate results to the client
  - For small datasets this process would be an acceptable option

The data is provided to the client with formatting to provide quick reference decision-making

![2017 Original Stock Results with Performance Timing](https://github.com/jkannis/stock-analysis/blob/main/Resources/2017_Original.png)

![2018 Original Stock Results with Performance Timing](https://github.com/jkannis/stock-analysis/blob/main/Resources/2018_Original.png)

- The refactored script used to process the stocks data used logic that looped through the dataset once, storing data in arrays for each ticker. After all data was processed, the arrays were looped through to write data to the sheet.
  - This process was successful in providing accurate results to the client
  - This logic was much more efficient and ran much more quickly than the original script making it a better choice for any size dataset

The data is provided to the client with formatting to provide quick reference decision-making

![2017 Refactored Stock Results with Performance Timing](https://github.com/jkannis/stock-analysis/blob/main/Resources/2017_Refactored.png)

![2018 Refactored Stock Results with Performance Timing](https://github.com/jkannis/stock-analysis/blob/main/Resources/2018_Original.png)

### Summary
#### Advantages and Disadvantages of Refactoring
The decision to refactor existing code is not a difficult one. The advantage of refactoring is the potential for increased efficiency in processing times. The disadvantages of refactoring are the time it takes to make modifications to existing code and the potential for introducing errors. When deciding whether or not to refactor, a couple of things should be considered. 
  1. How often will the script be executed
  2. How large are the datasets being used for the process

If the script is working accurately and will only be used occasionally, then the time required for refactoring may not be worth the effort. But if the script will be executed frequently, or the dataset being processed is very large, then taking the time to refactor would likely be beneficial.
#### Advantages and Disadvantages of Original vs. Refactored Scripts
Our client intends to run this analysis against large datasets of stock data with many years of information. While the original script was effective and accurate in providing the required information, the refactored script ran much more quickly and provided the same information. A possible disadvantage of the refactored script is it assumes the data is sorted by ticker. If it is provided with data that is not sorted in this way the results could be very unpredictable and messy, whereas the original script looped through the code looking for a specific ticker and could accumulate information regardless of the sort order of the data. 
