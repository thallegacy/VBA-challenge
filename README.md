VBA-Challenge - The VBA of Wall Street
Lynell Robinson
VBA scripting to analyze real stock market data. Data Includes:
•	The ticker symbol.
•	Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
•	The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
•	The total stock volume of the stock.
•	Bonus
o	The greatest percent change increase
o	The greatest percent change decrease
o	The great total stock volume

Method/Actions
•	Declared the worksheet variables
•	Set the worksheet variable
•	Created a loop to go through each sheet
•	Declared variables for stock value results within the worksheet loop
•	Declared loop assisting variables
•	Declared bonus variables stock value results
•	Set the initial counter to print yearly and percent change values
•	Set the initial stock result values
•	Defined a variable of the last row for iteration through each sheet
•	Created the iteration loop that will happen within in sheet to get the values
•	Created if statement to check when a ticker changes to a new value on each sheet
•	Print header rows for each sheet
•	Print bonus greatest columns headers for each sheet
•	Set the result values for Ticker, Yearchange, Totalstockvalue
•	Created if statement to check and solve for dividing by 0 for Percentchange
•	Print the values for Ticker, Yearlychange, Percentchange, Totalstockvalue
•	Created if statement for conditional formatting of yearlychange
•	Reset values to control next iterations at the change of the ticker point
•	Continue calculating value for Totalstockvalue when not a different ticker
•	Defined the range to look for the Greatest Percent Increase and its Ticker(Max of Percentchange)
•	Print Greatest Percent Increase and its Ticker
•	Defined the range to look for the Greatest Percent Decrease and its Ticker(Min of Percentchange)
•	Print Greatest Percent Decrease and its Ticker
•	Defined the range to look for the Greatest Total Stock Volume and its Ticker(Max of totalstockvalue)
•	Print Greatest Total Stock Value and its Ticker
•	Auto adjust columns
