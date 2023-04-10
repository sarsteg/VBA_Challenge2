Attribute VB_Name = "Module1"

''''''''''''''''''''''''''''''''''''''
''''' I. RUN FOR ALL WORKSHEETS  '''''
''''''''''''''''''''''''''''''''''''''

Sub RunAllWorksheets()

'Set a variable to WorkSheet
'This is to help shorten references to Worksheet
Dim ws As Worksheet

'The code below will keep the screen from updating, which can help the code run faster
Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        ws.Select
        ' Our actions are all saved in this macro
        ' If we wanted to have all of our steps in one macro...
        ' we would replace the Call command below with the code in the named macro
        Call RunSingleWorksheet(ws)

    Next ws

'Resetting to True for future
Application.ScreenUpdating = True

MsgBox "All done!"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''
''''' II. RUN FOR EACH INDIVIDUAL SHEET '''''
'''''''''''''''''''''''''''''''''''''''''''''

Sub RunSingleWorksheet()


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' A. Clean the data - WILL BE SKIPPING
    'These steps might be run if we had data that needed to be cleaned
    ' But our data appears to have already been cleaned and sorted
    ' These actions will not be taken

    ' 1. Check that ticker is only three characters, no spaces

    ' 2. Sort by date
        ' Currently the date is sorted as a text and not as a date
        ' Later, I'll covert this to look more like a date
        ' But for now, this needs to be a number

    ' 3. Sort the variables 
        ' decending date b/c we need to make sure that the earliest data is first and the latest date is last
        ' tickers are then sorted alphabetically so they are grouped in order


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' B. Conceptualizaion of the first table - DONE

    ' 0. Pre-work
        ' a. Setup variables to be used
            ' To remember the size of data types
            ' int (3 letters) < long (4 letters) < double (6 letters)

            ' i. Maintaining shortened Worksheet variable previous macro
                Dim ws As Worksheet
                Set ws = ThisWorkbook.ActiveSheet

            ' ii. Retrieval of Data
                ' Script loops through one yar of stock data and reads/stores all of the following values from each row
                ' Set type for each variable
                Dim ticker_symbol AS String 'Three letters
                Dim volume_of_stock AS Double 'Big number
                Dim open_price AS Double 'includes decimal
                Dim close_price AS Double 'includes decimal

                ' Set values based on first row of data, which are in a static location unless noted
                ' This setup is particularly valuable for open_price
                ' Because we will not set it when we start the loop
                ' We will only update it to the next ticker once we completed our current ticker
                ticker_symbol = ws.Cells(2, 1).Value 'Current ticker symbol
                volume_of_stock = 0 'Will capture when used
                open_price = ws.Cells(2, 3).Value 'Re-capturing at end of previous ticker
                close_price = 0 ' We need to capture this at last row of ticker

            ' iii. Variables for New Columns, Column Creation
                ' New columns created no the same worksheet with correct calculations
                ' New variables wil be calculated 
                ' Set type for each variable
                Dim total_stock_volume AS Double ' percent?
                Dim yearly_change AS Double ' money $?
                Dim percent_change AS Double ' percent again? 

                ' Set values to blank
                total_stock_volume = 0
                yearly_change = 0
                percent_change = 0

            ' iv. Last row for original data
                ' Create the variable
                Dim T1_last_row As Double
                ' Set calculation for variable
                T1_last_row = ws.Cells(Rows.Count,1).End(xlUp).Row
            
            ' v. Track row placement in new table
                ' Keep track of the row where we should place the next ticker in the summary table
                Dim T1_row_placement AS Double
                ' Start under the headers for the summary table
                T1_row_placement = 2

        ' b. Create headers
            'Headers for first table, starting in col I or 9
            ws.Cells(1,9) = "Ticker"
            ws.Cells(1,10) = "Yearly Change"
            ws.Cells(1,11) = "Percent Change"
            ws.Cells(1,12) = "Total Stock Volume"
   
    ' 1. Calculate summary values for each ticker

        ' Note: 
        ' Search through all the data row by row
        ' Determine if our current row is the last row of our current ticker
        ' We are doing this in order to determine the location of the last row of data to transition
        ' We will grabbing additional variables for calculations at the last row of current ticker

        ' Need to add variables for the current row and current ticker

        ' Read the row
        Dim i As Long
        For i = 2 To T1_last_row 

            ' a. If current row matches our current ticker: 
                If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then

                ' i. Capture this row's volume of stock
                    volume_of_stock = ws.Cells(i, 7).Value

                ' ii. Add current row to our running total
                    total_stock_volume = total_stock_volume + volume_of_stock
                
                ' iii. Will keep going to the next row below, Next i
 
            ' b. If current row does NOT match our current ticker:
                ElseIF ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

                ' i. Calculate values for the summary table
                
                    ' Close Price
                    ' Capture the closing price for the year and assign to variable
                    ' This will be used in calculations in the table
                    ' Closing price is in column F or 6
                    close_price = ws.Cells(i, 6).Value

                    ' Yearly Change
                    ' open_price was already captured
                    ' yearly_change equals closing_price minus open_price
                    yearly_change = close_price - open_price

                    ' Percent change
                    ' yearly_change created above
                    ' percent_change equals yearly_change divided by open_price
                    percent_change = (yearly_change / open_price)

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

                    ' Total Stock Volume
                    ' Capture this row's volume of stock
                    volume_of_stock = ws.Cells(i, 7).Value
                    ' Add in the currect row to the running total to update the variable 
                    total_stock_volume = total_stock_volume + volume_of_stock

                ' ii. Print in summary table

                    ' Our current ticker, col I
                    ws.Range("I" & T1_row_placement) = ticker_symbol

                    ' Yearly Change, col J
                    ws.Range("J" & T1_row_placement).NumberFormat = "0.00"
                    ws.Range("J" & T1_row_placement) = yearly_change

                    ' Percent Change, col K
                    ws.Range("K" & T1_row_placement).NumberFormat = "0.00%"
                    ws.Range("K" & T1_row_placement) = percent_change

                    ' Total Stock Volume, col L
                    ws.Range("L" & T1_row_placement) = total_stock_volume

                ' iii. Prep for next ticker
                    ' Now that the values have been used for our current ticker
                    ' Reset values for use in the next ticker

                    ' Move down one row in the summary table to place the next ticker
                        T1_row_placement = T1_row_placement + 1

                    ' Reset the running total
                        total_stock_volume = 0

                    ' Capture the next ticker symbol
                        ticker_symbol = ws.Cells(i + 1, 1).Value

                    ' Capture the next ticker's open price
                        open_price = ws.Cells(i + 1, 3).Value

                    ' Other variables will be reset or recalculated when last row of next ticker is found
                    ' volume_of_stock
                    ' close_price
                    ' yearly_change
                    ' percent_change
                    
            ' End this series of if statements
            End if

        Next i


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' C. Conceptualization of the Second Table

    ' 0. Pre-work
        ' a. Setup variables to be used in the second table
            'Double check that there are no duplicates

            ' i. Variables for calculations
                ' double check that all variables are unique from those used above

                ' Calculations for table
                ' Might need to consider how the variables are going to be rounded
                ' Possibly include rounding in the previous times this is calculated so that we have a stable number to reference each time
                ' Rather than a long number going out an unknown number of decimal places
                    Dim max_percent_change As Double
                    Dim min_percent_change As Double
                    Dim max_total_volume AS Double
                                                   
            ' ii. Determine lastrow  for table 2 that was just created
                ' Create variable
                Dim lastrowT2 As Long
                ' Calculate
                lastrowT2 = ws.Cells(Rows.Count,1).End(xlUp).Row

        ' b. Create headers for the table
            ' The example image leaves two extra columns, but I'm only going to separate by one column
            ' I think it looks cleaner with one column
            ' And is more consistent with the spaceing for the first table

            'Columns for summary table, starting in col 0 or 15
            ws.Cells(1, 15) = "Ticker"
            ws.Cells(1,16) = "Value"

            'Rows for summary table
            ws.Cells(2,14) = "Greatest % increase"
            ws.Cells(3,14) = "Greatest % decrease"
            ws.Cells(4,14) = "Greatest total volume"

    ' 1. Perform calculations
        ' Set range variables 
            Dim col_percent_change As Range
            Dim col_total_volume As Range

        ' Create range variables 
            Set col_percent_change = ws.Range("K2:K" & lastrowT2)
            Set col_total_volume = ws.Range("L2:L" & lastrowT2)
    
        ' a. Calculate the max percent change
            max_percent_change = Application.WorkSheetFunction.Max(col_percent_change)    

        ' b. Calculate the min percent change
            min_percent_change = Application.WorkSheetFunction.Min(col_percent_change)

        ' b. Calculate the max total volume
            max_total_volume = Application.WorkSheetFunction.Max(col_total_volume)

    ' 2. Find those numbers

        ' Note:
        ' Use lottery numbers activity as example of coding used
        ' Percent Change column is in col K or 11
        ' Total Stock Volume is in col L or 12
        ' Ticker is in col I or 9
        ' Tickers should be printed in col O or 15
        ' Values should be printed in col P or 16

        For x = 2 To lastrowT2 

            ' a. Find max 
                ' Look in col 11 for greatest % increase (max_percent_change)
                If ws.Cells(x, 11).Value = max_percent_change Then
                    
                    ' i. Include ticker in Summary Table, row 2, col 15
                        ws.Cells(2, 15) = ws.Cells(x, 9).Value
                    ' ii. Print max_percent_change in Table, row 2, col 16
                        ws.Range("P2").NumberFormat = "0.00%"
                        ws.Cells(2, 16) = max_percent_change

            ' b. Find min
                ' Look in col 11 for greatest % decrease (max_percent_change)
                ElseIF ws.Cells(x, 11).Value = min_percent_change Then

                    ' i. Include ticker in Summary Table, row 3, col 15
                        ws.Cells(3, 15) = ws.Cells(x, 9).Value
                    ' ii. Print min_percent_change in Table, row 3, col 16
                        ws.Range("P3").NumberFormat = "0.00%"
                        ws.Cells(3, 16) = min_percent_change

            ' c. Find max volume
                ' Look in col 12 for greatest total volume (max_total_volume)
                ElseIF ws.Cells(x, 12).Value = max_total_volume Then

                    ' i. Include ticker in Summary Table, row 4, col 15
                        ws.Cells(4, 15) = ws.Cells(x, 9).Value

                    ' ii. Print max_total_volume in Table, row 4, col 16
                        ws.Cells(4, 16) = max_total_volume

            ' End this series of If/Else
            End If
        
        ' Go to the next row to check values
        ' Continues until reaching lastrowT2 value
        Next x


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' D. Beautify the Tables

    ' 1. Adjust original table
    
        ' a. Match headers on original table to the other tables
            ' Remove <> surrounding names
                ws.Cells(1, 1) = "Ticker"
                ws.Cells(1, 2) = "Date"
                ws.Cells(1, 3) = "Open"
                ws.Cells(1, 4) = "High"
                ws.Cells(1, 5) = "Low"
                ws.Cells(1, 6) = "Close"
                ws.Cells(1, 7) = "Volume"  

        ' b. Format the date column as a date
            ' YYYY/MM/DD
            ' This is not for function for but aesthetics 

            ' Set a range for the column 
            ' We will use set instead of Dim because we want the column to be referred
            ' With Dim, we would be copying the values of the column into a new range
            ' Dim date_range As Range
            ' Set date_range = ws.Range("B2:B" & T1_last_row)

            ' Format new column
            'date_range.NumberFormat = "yyyy/mm/dd"

    ' 2. Formatting to look more like a table

        ' a. Original table
        ws.Range("A1:G1").Font.Bold = True

        ' b. Calculations table
        ws.Range("I1:L1").Font.Bold = True

        ' c. Summary table
        ws.Range("O1:P1").Font.Bold = True
        ws.Range("N2:N4").Font.Bold = True

    ' 3. Conditional formatting for Yearly Change column AND the percent change column

        ' Setup ranges for the columns
        ' col_percent_change was already created
        Dim col_yearly_change As Range
        Set col_yearly_change = ws.Range("J2:J" & lastrowT2)

        ' Create conditional formatting for yearly change column
        ' If value is > 0, fill with green; if value is < 0, fill with red, if value is = 0, fill with grey

        For each cell In col_yearly_change
            If cell.Value > 0 Then
                cell.Interior.ColorIndex = 4
            ElseIf cell.Value < 0 Then
                cell.Interior.ColorIndex = 3
            Else
                cell.Interior.ColorIndex = 15
            End If
        Next cell
        
        For each cell In col_percent_change
            If cell.Value > 0 Then
                cell.Interior.ColorIndex = 4
            ElseIf cell.Value < 0 Then
                cell.Interior.ColorIndex = 3
            Else
                cell.Interior.ColorIndex = 15
            End If
        Next cell

    ' 4. Resize the columns 
    
        ' a. All table columns show values
        ws.Range("A:G").Columns.AutoFit
        ws.Range("I:L").Columns.AutoFit
        ws.Range("N:P").Columns.AutoFit



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub