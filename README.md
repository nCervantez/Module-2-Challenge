# Module-2-Challenge
Stock Price Change Analysis Script
------------------------------------------------------------------------------------------------

Main Function
-----------------------------------------------------------------------------------

The scripts included in this repository will allow a user to take every stock listed in a file and see the yearly change, percent change and total stock volume for each stock placed into a table. This will allow for easy comparison of a stocks change from the yearly open to yearly close price of that stock, against every other stock listed in a file, for every worksheet.



Features
-----------------------------------------------------------------------------------
The codes included in this repository will take a list of stocks, their daily open and close prices, and compare the open price for the year to the close price and calculate the yearly change value, and percentage relative to the open price at the beginning of the year. The scripts will make a table and include the ticker, the yearly change, the percentage change and the total stock volume of the stock on the table, in a list format, for each stock included in the worksheet. 

The scripts also include conditional formatting to better visualize whether a stock had a positive or negative change from the initial year open price, in the form of green or red fill in a cell for positive and negative change respectively.

The scripts will also create a table that will list which stocks had the greatest percent increase, decrease, and which stock had the highest total volume for each year. The table will include the stock ticker and the value for each category in the table.

'ForEachWS.vb' will do as listed above for every worksheet in a file simultaneously.

Included in the repository are screenshots showing the code working across all sheets, and showing the full range of the table including every stock listed in each worksheet.

Additional Resources Used
-----------------------------------------------------------------------

https://stackoverflow.com/a/20648352/21964848

The linked resource above was used to apply number formats to specific cells, or ranges using proper VBA syntax.

\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

https://stackoverflow.com/a/4228480/21964848

The linked resource gave me the basis to create a boolean variable that would change for each new stock to capture the initial yearly opening price of the stock, without capturing any other opening price for that stock. This variable would reset to allow for this to function for each new stock encountered.

\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


https://officetuts.net/excel/vba/find-the-maximum-and-minimum-value-in-the-range-in-vba/

I used this resource to learn the VBA version of an excel function that would give me the maximum and minimum value of a range, and set them to variable I could call on in the code for multiple 'for loops'.

\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


https://learn.microsoft.com/en-us/office/vba/api/excel.colorindex

This resource provided a chart for the color index used in conditional formatting.


