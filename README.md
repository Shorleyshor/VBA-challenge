# VBA-challenge
Inserted a module for the multiple year stock excel file for the code.
created multiple Dims to declare the variables used in the script.
print the headings for the different rows containing values.
set initial total volume to 0 and initial table row to 2.
As the rows are to long, i set a lastrow function using the excel function formula.
created a for loop that loop through all ticker (using if, elseif and else statements) to retrieve ticker name, total volume, opening price at the begining of the year and closing price at the end of the year.
table row is increased by 1 for each loop and total volume is reset back to 0.
from the closing and opening price, i was able to calculate the yearly change and percentage change in price.
created a range.value to set the cells where the ticker, yearly change, percentage change, total stock volume, greatest % increase, greatest % decrease and greatest total volume are printed on the summary table.
set number format for how the numbers should appear in the table and both the percentage and the right decimal places were set.
set interior color index for values less than 0 (negative) to red and values greated than 0 (positive) to green. using a for loop.
The bonus question was done using both excel functions and for loop to get both minimum and maximum numbers to retrieve the greatest % increase, greatest % decrease and greatest total volume. And the values were formated so they appear in the correct % and decimals.
A WORKSHEET code was created using excel functions and for loop to set and make appropriate adjustment so the VBA script can run on every worksheet just by running the VBA script once.