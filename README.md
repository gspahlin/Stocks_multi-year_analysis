# Stocks_multi-year_analysis

Project Description:

The goal of this project was to write an excel macro that does various calculations on a spreadsheet with daily data on several thousand stocks from several years. Running the macro will calculate the difference in the first (non-zero) opening value and last closing value for each stock, calculate the sum of all the volumes for each stock, and display this number next to the ticker in a new table to the left. The percentages are color coded - red for zero or negative values, and green for positive value. A new table will then be generated to display the maximum and minimum values of volume and percent change for the year. The macro is also designed to analyze all sheets in range, so there is no need to run it more than once. 

Here are my descriptions of the files provided in this repository to document this project:

Stock_calc_hard.bas  - This is a visual basic file that contains the code of the finished macro plus comments. It can be opened with a text editor to inspect my solution

Multiple_year_stock_data_Hard_Challenge_full2.xlsm - this is a file that contains the stock data, which has been been analyzed using the macro

Multiple_year_stock_data_macro_ready.xlsm - this is a version of the file above with the macro loaded, but saved prior to running it. This macro can be run to confirm that the changes in the file above came from the macro, and to time it if desired. To run the macro go to "modules" in the developer tab, and look for "Stock_Calc_Hard" this is the name of the macro. 

alphabetical_testing_3 - this is a smaller sheet that contains an unfinished version of the macro. I included it because tab A shows some manual calculations I did to test whether the macro was calculating desired quantities correctly. 

