Attribute VB_Name = "Module1"
Sub stock_analysis()
'loop through all worksheets
For Each ws In Worksheets
    'variable to hold the ticker name
    Dim Ticker_Name As String
    Ticker_Name = " "
    'variable for holding the total per tick name
    Dim ticker_total As Double
    Ticker_Name = 0
    'variable for total stock volume
    Dim Total_stock_volume As Double
    Total_stock_volume = 0
    'variables for initial assignment
    Dim Open_Price As Double
    Open_Price = ws.Cells(2, 3).Value
    Dim Close_Price As Double
    Close_Price = 0
    Dim change_in_price As Double
    change_in_price = 0
    Dim change_in_percent As Double
    change_in_percent = 0
    '_______________________________________________________
    'Calculations for a single worksheet
    'will change once it works
    'Got code to work with one page on (7/3/2020), now to make it work with more worksheets
    'keeping track of the location for each ticker name in a summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    'setting last row loop
    Dim last_Row As Long
    Dim i As Long
    last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'set titles for summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    'Titles for challenge variables
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    'loop through all tickers
        For i = 2 To last_Row
    'run a check to see if the we're in the same ticker, if not:
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'Set ticker name
        Ticker_Name = ws.Cells(i, 1).Value
        'tell code where close price are
        Close_Price = ws.Cells(i, 6).Value
        'calculate yearly change
        change_in_price = (Close_Price - Open_Price)
        'add to stock volume
         Total_stock_volume = Total_stock_volume + ws.Cells(i, 7).Value
         'correct divistion by zero bug
         If Open_Price = 0 Then
            change_in_percent = 0
         Else
            change_in_percent = change_in_price / Open_Price
         End If
         'print the ticker symbol in the summary table in Column I
         ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
         'Print the change in price in summary table under Column J
         ws.Range("J" & Summary_Table_Row).Value = change_in_price
         'Print the Percentage Change in summary Table under Column K
         ws.Range("K" & Summary_Table_Row).Value = (Str(change_in_percent))
         ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
         'Print the total stock volume in summart table under Column L
         ws.Range("L" & Summary_Table_Row).Value = Total_stock_volume
         'Add one to the summary row
         Summary_Table_Row = Summary_Table_Row + 1
         'reset variables for new tickers
         Ticker_Name = 0
         Total_stock_volume = 0
         change_in_price = 0
         Close_Price = 0
         Open_Price = ws.Cells(i + 1, 3)
      Else
         'add to the stock total
         Total_stock_volume = Total_stock_volume + ws.Cells(i, 7).Value
        End If
    Next i
    
'code for summary table last row to make bonus section easier
    last_row_summary = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        'Conditonal formatting for colors
         For i = 2 To last_row_summary
          If ws.Cells(i, 10) >= 0 Then
             ws.Cells(i, 10).Interior.ColorIndex = 10
          Else
             ws.Cells(i, 10).Interior.ColorIndex = 3
          End If
    Next i





'Bonus sections for Greatest % increase, decrease, and total volume
'To find the greatest increase we have to use the change in percent for each ticker
    'First find the max
    For i = 2 To last_row_summary
    'application.WorksheetFunction.max reference: https://bit.ly/2ArO7PN
     If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & last_row_summary)) Then
        ws.Range("P2").Value = ws.Cells(i, 9).Value
        ws.Range("Q2").Value = ws.Cells(i, 11).Value
        ws.Range("Q2").NumberFormat = "0.00%" 'format for percent reading
    'application.WorksheetFunction.min reference: https://bit.ly/2ArO7PN
     ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & last_row_summary)) Then
        ws.Range("P3").Value = ws.Cells(i, 9).Value
        ws.Range("Q3").Value = ws.Cells(i, 11).Value
        ws.Range("Q3").NumberFormat = "0.00%" 'format for percent reading
    'code for the volume of these stocks
     ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & last_row_summary)) Then
        ws.Range("P4").Value = ws.Cells(i, 9).Value
        ws.Range("Q4").Value = ws.Cells(i, 12).Value
    End If
    Next i
    'to avoid any ####### values use the autofit function
    ws.Columns("I:L").EntireColumn.AutoFit
    ws.Columns("O:Q").EntireColumn.AutoFit
    Next ws
End Sub
