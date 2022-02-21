# VBA-Homework

#We set the variables first, that is:

#The Ticker as a String variable since it is a categorical variable

#Yearly Change (Change) and Percentage (PercentChange) to be as Double to ensure that they have 2 decimal places

#And finally, Stock Volume (Stock_Vol) as a LongLong because it would overflow if this does not apply.

#We then set the Yearly Change, Percentage Change and Stock Volume to zero to start the process with nothing there which eventually prevents errors when running through the loops.

#A Summary Table (Summary_Table_Row) is set to ensure that the 4 variables are in a table in the final result.

#Another 4 variables are set, which include OpenP (the Opening Price), CloseP (the Closing Price), Start (as the starting variable) and Loc (the value we want to locate for which is the starting variable)

#The headings of the summary table per column are then set.

#We then set the Opening Price and Closing Price to zero to start these values fresh.

#The starting value is then equal to 2 to locate the opening price.

#For the location of the entire table:

#If the tickers do not match between rows connected to each other, we set the location of the ticker and the stock volume (first and last columns of the summary table)

#We also set the variables of opening price and closing price.

#After that, the opening price is to be set as the first opening price of the year.

#Formulas are set for yearly change and percentage change.

#The location of the four variables (Ticker, Yearly Change, Percentage Change and Stock Volume) are set in the Summary Table. We ensure that the Percentage Change is in a percentage value.

#The variables are then reset to start the entire loop again.

#The colours of Yearly Change are set to red if less than 0 or green if 0 or above.
