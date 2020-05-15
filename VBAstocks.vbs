Attribute VB_Name = "Module1"
Sub stock_code()
 
For Each ws In Worksheets
 'declare all the variables to be used
    Dim ticker As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim volume As LongLong
    Dim RowNum As Integer
    Dim openvalue As Double
    Dim closevalue As Double
    Dim lastrow As Long
    Dim max_volume As LongLong
    Dim max_change As Double
    Dim min_change As Double

   'set initial values
        volume = 0
        RowNum = 2
        max_volume = 0
        max_change = 0
        min_change = 0
    'lable the headers for summary table
    ws.Range("I1").Value = "ticker"
    ws.Range("J1").Value = "yearly change"
    ws.Range("K1").Value = "percent change"
    ws.Range("L1").Value = "volume"
    ws.Range("P1").Value = "ticker"
    ws.Range("Q1").Value = "value"
    ws.Range("O2").Value = "Greatest Percent Increase"
    ws.Range("O3").Value = "Greatest Percent Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'define the last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'define open price-gets reset in loop
    openvalue = ws.Cells(2, 3).Value
    
    'begin the for loop
    For i = 2 To lastrow
    'checks if the ticker value are different
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    'assigns a value for the ticker variable
            ticker = ws.Cells(i, 1).Value
    'assigns value for the close value and the yearly change, uses these to calculate percent change
            closevalue = ws.Cells(i, 6).Value
            yearly_change = closevalue - openvalue
    'check to make sure the denominator is not 0
                If openvalue <> 0 Then
                    percent_change = yearly_change / openvalue
                End If
    'adds last row volume to total volume
            volume = volume + ws.Cells(i, 7).Value
    'populates the data to the summary table
            ws.Range("I" & RowNum).Value = ticker
            ws.Range("J" & RowNum).Value = yearly_change
            ws.Range("K" & RowNum).Value = percent_change
            ws.Range("L" & RowNum).Value = volume
    'populates data into challenge summary table
            If percent_change > max_change Then
                ws.Range("P2").Value = ticker
                ws.Range("Q2").Value = percent_change
                max_change = yearly_change
            ElseIf percent_change < min_change Then
                ws.Range("P3").Value = ticker
                ws.Range("Q3").Value = yearly_change
                min_change = percent_change
            End If
            
            If volume > max_volume Then
                ws.Range("P4").Value = ticker
                ws.Range("Q4").Value = volume
                max_volume = volume
            End If
    'readies data for next loop
            RowNum = RowNum + 1
            volume = 0
            openvalue = ws.Cells(i + 1, 3).Value
            
       ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
       'adds volume for total volume and defines the opening value
            volume = volume + ws.Range("G" & i).Value
            
   End If
   Next i
   
   For i = 2 To lastrow
   
   'conditional formatting for coloring in the yearly change column
    If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
        
    End If
    Next i
'format for percentage
    ws.Range("K2:K" & lastrow).NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
Next ws

End Sub

