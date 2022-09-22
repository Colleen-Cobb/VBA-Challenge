Attribute VB_Name = "Module1"
Sub forloopStock_Market_Analysis()

    'Declare Current as worksheet object variable
Dim ws As Worksheet

Dim WorksheetName As String

    'Loop through all worksheets
For Each ws In Worksheets

   'Determine the Last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Grab the worksheet name
WorksheetName = ws.Name

    'Add the word Ticker to the (1,9) cell
ws.Cells(1, 9).Value = "Ticker"

    'Add Yearly Change to the cell (1,10)
ws.Cells(1, 10).Value = "Yearly Change"

    'Add Percent Change to the cell (1,11)
ws.Cells(1, 11).Value = "Percent Change"

    'Add Total Stock Volume to the cell (1,12)
ws.Cells(1, 12).Value = "Total Stock Volume"

    'Set a variable for opening price at the beginning of the year
Dim Opening_Price As Double

    'Set a variable for closing price at the end of the year
Dim Closing_Price As Double

    'Set a variable for percent change
Dim Percent_Change As Double

    'Keep track of the location for each ticker symbol in the sumary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

    'Assign an integer for total stock volume count to start at
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0

    'Set a variable for the loop to start moving to the next ticker symbol and assign a value
Dim Last_Ticker_Symbol
Last_Ticker_Symbol = 1

    



'*********************************************
'LOOP THROUGH DATA TO OUTPUT THE TICKER SYMBOL, YEARLY CHANGE, PERCENT CHANGE, and TOTAL STOCK VOLUME
'*********************************************

    'Determine the Last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


    'Set an initial varibale for holding the ticker symbol
Dim Ticker_Symbol As String

    'Create a script that loops through all the stocks in each worksheet and outputs the necessary values
    
For i = 2 To LastRow

    'Check if we are wtihin the same stock name, if we are not...
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    'Set Ticker Symbol
Ticker_Symbol = ws.Cells(i, 1).Value

    'Allow the ticker symbol to move to the next ticker symbol
Last_Ticker_Symbol = Last_Ticker_Symbol + 1

    'Get the opening price and the closing price of the year
Opening_Price = ws.Cells(Last_Ticker_Symbol, 3).Value
Closing_Price = ws.Cells(i, 6).Value

    'Create an inner loop to determine the total stock volume for each ticker symbol
For j = Last_Ticker_Symbol To i

    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(j, 7).Value

Next j

If Opening_Price = 0 Then

    Percent_Change = Closing_Price

Else
    Yearly_Change = Closing_Price - Opening_Price
    Percent_Change = Yearly_Change / Opening_Price
    
End If

    'Output the values in the summary table
ws.Cells(Summary_Table_Row, 9).Value = Ticker_Symbol
ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
ws.Cells(Summary_Table_Row, 11).Value = Percent_Change

    'Change the number format in percnet change cell to percentage
ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"

 'Use condiitional formatting to highlight positive changes in green and negative changes in red
If ws.Cells(Summary_Table_Row, 10).Value < 0 Then
    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
    
Else: ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4

End If

    
    'Ensure Total stock volume is going to the correct cell range
ws.Cells(Summary_Table_Row, 12).Value = Total_Stock_Volume

    'Move to the next row in the summary row table once one is completed
Summary_Table_Row = Summary_Table_Row + 1

    'Change the variables back to 0 to reset counting
Total_Stock_Volume = 0
Yearly_Change = 0
Percent_Change = 0


    'Move the row to Last Ticker Symbol
Last_Ticker_Symbol = i

 


End If

Next i


Next ws

End Sub
