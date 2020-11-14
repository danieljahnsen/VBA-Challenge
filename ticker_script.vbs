Sub ticker_tracker():

'Loop through the sheets in the file:
For Each ws In Worksheets:
    
    'Find the last row in the active worksheet
    lastrow = Int(ws.Cells(Rows.count, 1).End(xlUp).Row)

    'Insert the Headers for each Worksheet
    ws.Cells(1, 9).Value() = "Ticker"
    ws.Cells(1, 10).Value() = "Yearly Change"
    ws.Cells(1, 11).Value() = "Percent Change"
    ws.Cells(1, 12).Value() = "Total Stock Volume"
    
    'Insert the Bonus Table
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
    'Inserting Bonus Table Default Values
    ws.Range("P2").Value = 0
    ws.Range("P3").Value = 0
    ws.Range("P4").Value = 0
    
    Dim name As String

    ' Set an initial variable for the volume
    Dim volume As Double
    volume = 0
    
    'Set an initial variable for the start price of a stock
    Dim start As Double
    start = ws.Cells(2, 3).Value
    
    'Set an initial variable for the final price of a stock
    Dim final As Double
    final = 0
    
    ' Keep track of the location for each Ticker
    Dim count As Integer
    count = 2

    ' Loop through all tickers
    For i = 2 To lastrow

        ' Check if we are still within the same ticker
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        ' Set the Ticker name
        name = ws.Cells(i, 1).Value
        
        'Set the final value of the stock
        final = ws.Cells(i, 6).Value

      ' Add to the volume
        volume = volume + ws.Cells(i, 7).Value

      ' Print the ticker in the summary table
        ws.Range("I" & count).Value = name

      ' Print the volume to the Summary Table
        ws.Range("L" & count).Value = volume
        
        'Calculate the yearly change and print to the table
        ws.Range("J" & count).Value = (final - start)
        
        'Formats the yearly change to green or red
        If (final - start) > 0 Then
            ws.Range("J" & count).Interior.ColorIndex = 4
        Else
            ws.Range("J" & count).Interior.ColorIndex = 3
        End If
        
        'Calculate the yearly change and print to the table and set the format to percentage
        'Only calcs if start is not zero
        If start <> 0 Then
            ws.Range("K" & count).Value = (final - start) / start
            ws.Range("K" & count).NumberFormat = "0.00%"

        End If

      ' Add one to the summary table row
        count = count + 1
      
      ' Reset the volume
        volume = 0
        
        'Sets the next start price
        start = ws.Cells(i + 1, 3).Value

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the volume
        volume = volume + ws.Cells(i, 7).Value
        

    End If
  Next i
  
    'Bonus
    'Assign bonus variable and set default values
    Dim g_increase, g_decrease, gvol As Double
    g_increase = 0
    g_decrease = 0
    gvol = 0
    
    'Formats the cells
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"


    'Loop through summary table
    For x = 2 To count:
        If ws.Cells(x, 11).Value > g_increase Then
            ws.Range("P2").Value = g_increase
            ws.Range("O2").Value = ws.Cells(x, 9).Value
            g_increase = ws.Cells(x, 11).Value
        ElseIf ws.Cells(x, 11).Value < g_decrease Then
            ws.Range("P3").Value = g_decrease
            ws.Range("O3").Value = ws.Cells(x, 9).Value
            g_decrease = ws.Cells(x, 11).Value
        ElseIf ws.Cells(x, 12).Value > gvol Then
            ws.Range("P4").Value = gvol
            ws.Range("O4").Value = ws.Cells(x, 9).Value
            gvol = ws.Cells(x, 12).Value
        End If
    Next x
    
Next ws

MsgBox ("All done")

End Sub