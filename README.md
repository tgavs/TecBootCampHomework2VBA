# TecBootCampHomework2VBA
VBA scripts to analyse stock prices and volume


-----------Quick Start----------

To Run the project, execute the Sub "main()" from vba

The sub main calls the following functions:

    Call List_and_volume(ws) 'creates a unique list and sets the total volume traded for each unique ticker

    Call Year_change(ws) 'calculates the yearly price change and pct_change

    Call Greatest_Changes(ws) 'calculate the greatest price changes

When its done you will receive a msgbox

---------Files----------------

alphabtical_testing-TGS2.xlsm -> ready to run excel file with the all the Sub intalled

Multiple_year_stock_data - TGS.xlsm -> this is the file with all the data and the Sub already run

Unite1-Assignment2-Report-vba.docx -> this file contains the screenshots of the analysis from the Multiple_year_stock_data file

--------Scripts------

List_and_volume.vbs -> creates a unique tickers list at the same time calculates the total volume traded optimizing the resources

Yrl_Change.vbs -> this sub calculates the yearly change for every stock

Greatest_changes.vbs  -> calculates the greatest changes among the stocks

-----Additional Scripts----
Another way to create the unique tickers list and total volume. Not too fast but using straightforward what we have seen in class.

This subs are not necesary to run the main

Lista_unica.vbs 

Total_volume.vbs



