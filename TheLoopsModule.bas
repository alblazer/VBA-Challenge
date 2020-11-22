Attribute VB_Name = "Module1"
Sub theloops2()
    Dim ws As Worksheet

    For Each ws In Worksheets
        ' Finding the last row of the combined sheet
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
        ' Finding the last ws row and returning the # of rows w/o the header
        lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1
        ' Copy the contents of each state sheet into the combined sheet
        'ws.Range("A" & lastRow & ":G" & ((lastRowState - 1) + lastRow)).Value = ws.Range("A2:G" & (lastRowState + 1)).Value
    Next ws
    'Copy the headers from sheet 1
    'ws.Range("A1:G1").Value = Sheets(2).Range("A1:G1").Value
    
    ' Autofit to display data
    'ws.Columns("A:K").AutoFit
    'Initialize the ticker, stock volumes, open and close years, as well as percentages
    Dim tickname As String
    Dim tickvolume As Double
    tickvolume = 0
    Dim stock_volume As Double
    stock_volume = 0
    Dim openyear As Double
    openyear = 0
    Dim closeyear As Double
    closeyear = 0
    Dim change As Double
    change = 0
    Dim percentage As Double
    percentage = 0
    'As well as the row counter
    Dim rowvalue As Integer
    rowvalue = 2
    'Set headers for the Ticker names and total stock volumes
    Range("H1").Value = "Tickername"
    Range("I1").Value = "total_stock_volume"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    openyear = Cells(2, 3).Value
    'Iterate through all the rows
    For i = 2 To lastRow
        'When the current Ticker isn't the same as the next one then
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'Set the ticker name
            tickname = Cells(i, 1).Value
            'Set the stock volume
            stock_volume = stock_volume + Cells(i, 7).Value
            'On row H, print the tickname
            Range("H" & rowvalue).Value = tickname
            'On row I, print the added stock volume
            Range("I" & rowvalue).Value = stock_volume
            'Increase the row counter and reset the stock colume counter
            rowvalue = rowvalue + 1
            stock_volume = 0
            'Set the current row's closing values for the year
            closeyear = Cells(i, 6).Value
            'Calculate the change
            change = closeyear - openyear
            'Set the change on Column J
            Range("J" & rowvalue - 1).Value = change
            'Find the percentage change for the year
            If openyear = 0 Then
                percentage = 0
            Else
                percentage = change / openyear
            End If
            'Set the result under Column K and change it's format to %
            Range("K" & rowvalue - 1).Value = percentage
            Range("K" & rowvalue).NumberFormat = "0.00%"
            'Fill "Yearly Change", i.e. Delta_Price with Green and Red colors
            If (change > 0) Then
                'Fill column with GREEN color - good
                Range("J" & rowvalue - 1).Interior.ColorIndex = 4
            ElseIf (change <= 0) Then
                'Fill column with RED color - bad
                Range("J" & rowvalue - 1).Interior.ColorIndex = 3
            End If
            
            'Reset the stock change and closing year for the next row
            change = 0
            closeyear = 0
            'Increment the opening year value by one row
            openyear = Cells(i + 1, 3).Value
            
        Else 'If the ticker name is the same, add their stock volume amount
            stock_volume = stock_volume + Cells(i, 7).Value
        End If
    Next i
        

End Sub

