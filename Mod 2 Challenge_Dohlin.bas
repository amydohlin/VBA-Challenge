Attribute VB_Name = "Module1"
Sub multi_year_stock():

Dim ws As Worksheet

For Each ws In Worksheets

    'variable containing ticker
    Dim Ticker As String
    Dim TickerRow As Long
    TickerRow = 1
    
    'variable containing changes in ticker thru out year, this is for the Yearly Change
    Dim Tick_Change As Double
    Tick_Change = 0
    Dim closep As Double    'get data from the <close> column
    Dim openp As Double     'get data from the <open> column
    
    'variable containing percent of change in ticker
    Dim Percent_Change As Double
    Percent_Change = 0

    'total volume of each ticker. Use "double" to hold values containing decimals
    Dim Total_Vol As Double
    Total_Vol = 0
    
    'summary location for each ticker type
    'headers for table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"

    'greatest hits table (% inc, % dec, total vol)
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        'using a For loop to read thru column A for tickers and pulling out each unique ticker
        'starting and ending row, using i for rows and j for columns
        Dim i As Long
        Dim j As Integer
        Dim Summary As Integer
        Summary = 2
        
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    For i = 2 To LastRow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
        
        TickerRow = TickerRow + 1
        Ticker = ws.Cells(i, 1).Value
        ws.Cells(TickerRow, 9).Value = Ticker
        
        ws.Range("I" & Summary).Value = Ticker
        
        'total volume for each ticker
        Total_Vol = Total_Vol + ws.Cells(i, 7).Value
        ws.Range("L" & Summary).Value = Total_Vol
        Summary = Summary + 1
        
        'yearly change
        closep = ws.Cells(i, 6).Value
        openp = ws.Cells(i, 3).Value
        Tick_Change = closep - openp
        ws.Range("J" & Summary).Value = Tick_Change
            'ISSUE GETTING YEARLY CHANGE TO PRINT FOR FIRST TICKER IN EACH SHEET
            'TRIED closep = ws.Cells(i+1, 6).Value AND openp = ws.Cells(i+1, 3).Value, DID NOT WORK

            ElseIf openp <> 0 Then
            Percent_Change = (Tick_Change / openp) * 100
            ws.Range("K" & Summary).Value = Percent_Change
            'DON'T UNDERSTAND THIS PART

    Else
    
        Total_Vol = Total_Vol + ws.Cells(i, 7).Value
        
        End If
        
    Next i
    
Next ws

MsgBox ("boop boop")

        


End Sub
