Attribute VB_Name = "Module1"

' I. For each worksheet (Run this for each sheet)

' II. Clean the data
    ' 1. Sort the variables 
        ' decending date b/c we need to make sure that the earliest data is first and the latest date is last
        ' tickers are then sorted alphabetically so they are grouped in order


    ' 2. Format the date column as a date
        ' Currently the date is sorted as a text and not as a date
        ' Insert a new column so that we can maintain data integrity and avoid writting over original data
        ' Hide the old date since it is not useful

'III. Conceptualization of the first table

' 0. Pre-work
    ' setup variables to be used
    ' include lastrow variable
    ' include the open value for the year, as it will show up in a consistent location each time at the start of the sheet

' 1. Look at the first cell for the group, grab the opening data for that column
    ' grab the name of the ticker

    ' 2. Should we keep going? 
        ' Look at the name in column A and determine if the cell below matches
        ' We are doing this in order to determine the location of the last row of data
        ' where we will grabbing additional variables for calculations

        ' a. If the cell matches: 
            'add up to our running totals for the total stock volume

        ' b. If the cell does NOT match:

            ' pull the name of the ticker from the row we are on and
            ' print this in the table where it should be (if not already done)
            ' set this name from the cell to be the ticker
            
            ' capture the closing price for the year and assign to variable
            ' this will be used in calculations in the table

            ' Yearly Change
            ' update yearly change variable = closing for the year minus openning for the year
            ' print yearly change variable in cell (moves with the rows)

            ' Percent change
            ' (New Price - Old Price) / Old Price x 100
            ' Divide yearly change variable by the closing price
            ' assign to variable percent change
            ' print percent change variable in cell (moves with the rows)

            ' Total Stock Volume
            ' add in the currect row to the running total to update the variable 
            ' print total stock volume variable in the cell (moves with the rows)

            ' now that the open value has been used in the previous calculation
            ' set up the open value to be the next set that we will be reviewing

            ' reset counter?

' IV. Conceptualization of the Second Table

' 0. Pre-work
    ' set all the variables that are going to be used in the second half
    ' double check that all variables are unique from those used above
    ' include a lastrow variable for the table that we will be referencing

    ' for use in all the all the calculations below
    ' create a range to reference for the percent change column
    ' create a range to reference for the total stock volume

    ' 1. Calculate the max
        '=max(range percent change)    

    ' 2. Calculate the min
        ' min(range percent change)
        ' index(match) to get the ticker

    ' 3. Calculate the highest total volume
        ' max(range total volume)

' V. Beautify the Tables

    ' 1. Create the headers

    ' 2. Formatting, such as bold text

    ' 3. Conditional formatting for Yearly Change column AND the percent change column
        ' if the value is > 0, then setup as green
        ' if the value is < 0, then setup as red
        ' if the value is = 0, then setup as grey

Sub RunAllWorksheets()

''''''''''''''''''''''''''''''''''''''
''''' I. RUN FOR ALL WORKSHEETS  '''''
''''''''''''''''''''''''''''''''''''''

'!!!!!LOOPING ACROSS WORKSHEET!!!!!

'Set a variable to WorkSheet
'This is to help shorten references to Worksheet
Dim ws As Worksheet

'The code below will keep the screen from updating, which can help the code run faster
'Application.ScreenUpdating = TRUE

For Each ws In wb.Worksheets
    ws.Select
    ' Our actions are all saved in this macro
    ' If we wanted to have all of our steps in one macro...
    ' we would replace the Call command below with the code in the named macro
    Call RunSingleWorksheet

Next ws

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub RunSingleWorksheet()


''''''''''''''''''''''''''''
''''' SETUP VARIABLES  '''''
''''''''''''''''''''''''''''

' Maintaining the variable assignment from previous macro
Dim ws As Worksheet

'Retrieve variables from existing table and set type 
Dim ticker_symbol AS String
Dim volume_of_stock AS Long
Dim open_price AS Double
Dim close_price AS Double

'Set values to blank
ticker_symbol = Cells(2, 2) 'Starts at static location 
volume_of_stock = 
open_price = Cells(2, 3) 'Stats at static location
close_price = 0 'We need to capture this

'New variables to calculate and set type
Dim total_stock_volume AS Double 'percent?
Dim yearly_change AS Double 'money $?
Dim percent_change AS Double 'percent again?

'Set values to blank
total_stock_volume = 0
yearly_change = 0
percent_change = 0

'New summary variables to calculate and set type
Dim greatest_increase AS Double 'percent?
Dim greatest_decrease AS Double 'percent?
Dim greatest_total_vol AS Double 

'Set values to blank
greatest_increase = 0
greatest_decrease = 0
greatest_total_vol = 0

'Determine when the lastrow is for the sheet
lastrow = Cells(Rows.Count,1).End(x1Up).Rows

'''''''''''''''''''''''''''''''''''''
''''' SETUP STRUCTURE OF TABLES '''''
'''''''''''''''''''''''''''''''''''''

'Headers for first table, starting in col I or 9
Cells(1,9) = "Ticker"
Cells(1,10) = "Yearly Change"
Cells(1,11) = "Percent Change"
Cells(1,12) = "Total Stock Volume"


'''''''''''''''''''''''''''''''
''''' CREATE THE DATA SET ''''' 
'''''''''''''''''''''''''''''''

















'Only run this while the row is not Empty
'While NOT IsEmpty(Cells(r, 1)) <remove b/c using last row counter

    'Read the row
    For i = 2 To 10000000000

        'Read the columns
        For j = 1 To 8

            '!!!!! RETREVAL OF DATA !!!!!
            'Set the variables for this row to be used in calculations

                'ticker symbol is in col A or 1
                ticker_symbol = Cells(i, 1)            

                'volume of stock is in col G or 7
                volume_of_stock = Cells(i, 7)

                'open price is in col C or 3
                open_price = Cells(i, 3)

                'close price is in col F or 6
                close_price = Cells(i, 6)
            
            '!!!!! COLUMN CREATION !!!!!
            'Create adjacent table

                'ticker symbol is going to col I or 9
                Cells(i, 9) = ticker_symbol

                'yearly change $ must be calculated
                'and is going to 10
                yearly_change = 'insert calculation
                Cells(i,10) = yearly_change

                'percent change must be calculated
                'and is going to col 11
                percent_change = 'insert calculation
                Cells(i,11) = percent_change

                'total stock volume must be calculated
                'and is going to col 12
                total_stock_volume = 'insert calculation
                Cels(i,12)=total_stock_volume

        'We will start back at j=1 to be in column A
        'So NO Next j, this would cause the columns to shift over each time by 1  
        'j will be reset to 1 again if the loop runs again 
    
    'Go to the next row
    Next i

'Go to the next row to check if it is blank
'r = r + 1 < taking this out and using the last row counter


''''''''''''''''''''''''''''''''''
''''' CONDITIONAL FORMATTING '''''
''''''''''''''''''''''''''''''''''

'!!!!!CONDITIONAL FORMATTING!!!!!



'''''''''''''''''''''''''''''''''''''''''''''
''''' SETUP STRUCTURE FOR SUMMARY TABLE '''''
'''''''''''''''''''''''''''''''''''''''''''''

'The example leaves two extra columns, but I'm only going to separate by one column
'I think it looks cleaner with one column
'Headers for summary table, starting in col 0 or 15
Cells(1, 15) = "Ticker"
Cells(1,16) = "Value"

'Rows for summary table
Cells(2,14) = "Greatest % increase"
Cells(3,14) = "Greatest % decrease"
Cells(4,14) = "Greatest total volume"


''''''''''''''''''''''''''''''''
''''' CREATE SUMMARY TABLE '''''
'''''''''''''''''''''''''''''''' 

'!!!!!CALCULATED VALUES!!!!!

'Find values
greatest_increase = 'calculation
greatest_decrease = 'calculation
greatest_total_vol = 'calculation

'Keep searching until nothing to search
While NOT IsEmpty(Cells(s, 1))

    'Include ticker in Summary Table
    For x = 2 To 10000000000

        'greatest % increase for sheet
        If Cells()

        'greatest % decrease

        'greatest total volume

        'Match-up ticker to pull remaining information to table

Next s

''''''''''''''''''''''''''''
''''' BEAUTIFUL TABLES '''''
''''''''''''''''''''''''''''

'Setup the columns so that all the text is visible
'Autosize to the width of the text


'Bold for the headers on the tables
'All of row 1 to bold from a to p

'Add bold for the row headers on summary table


'Resize columns H and M so they are the same size, but not as big as the default


