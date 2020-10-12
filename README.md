# VBA-Challenge

Just a note:

I added all parts of the HW assignment (the main assignment and the 2x challenges) to one .bas file.

## Challenge Details

This challenge was to create a VBA script that would take an excel file containing stock information for multiple stocks from the years 2014 to 2016 and output each stock's:

* yearly change from opening price at the beginning of a given year to the closing price at the end of that year
* percent change from opening price at the beginning of a given year to the closing price at the end of that year
* total stock volume of the stock

Next, conditional formatting was used to highlight yearly change in red if it was negative and green if it was positive.

As an additional challenge, the VBA script was modified to also return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". 

The VBA script is set to run on every worksheet (each year) by running it one time. 

See example output below for the year of 2014:

![2014 Stock Data](Multiple_year_stock_data_2014.png)

## Included Files

In this repository, the following files are included:
* Main BAS file - the VBA script needed to run through each excel worksheet and perform necessary calculations
* Snapshot of output for each year (2014, 2015, 2016)
