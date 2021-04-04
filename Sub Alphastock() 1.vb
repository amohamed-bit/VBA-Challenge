Sub Alphastock()
' LOOP THROUGH ALL SHEETS
For Each ws In Worksheets
'declaring variables'
Dim ticker As String
Dim opener As Double
Dim closer As Double
Dim total_vol As Double
Dim yearly_chn As Double
Dim percent_chn As Double
'initialize total volume'
total_vol = 0
'intiailizing summary table'
Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
' Determine the Last Row
   LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'column names for output table'
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Volume"

 ' Loop through all
  For i = 2 To LastRow

    ' Check if we are still within the same Ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    ' Set the ticker
      ticker = ws.Cells(i, 1).Value

      ' Add to the Total volume
      total_vol = total_vol + ws.Cells(i, 7).Value

      ' Print the ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = ticker

      ' Print the total volume to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = total_vol
      'set the closing price'
      closer = ws.Cells(i, 6).Value
      'set opening price'
      opener = ws.Cells(i, 3).Value
      
      'calculating yearly change'
      yearly_chn = closer - opener
      'calculating Percent change'
      If opener = 0 Then
      percent_chn = 0
      Else
      percent_chn = yearly_chn / opener
      End If
      If percent_chn > 0 Then
      ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 4
      Else
      ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 3
      End If
      
      ' Print in the Summary Table'
      ws.Range("J" & Summary_Table_Row).Value = yearly_chn
      ' Print in the Summary Table'
       ws.Range("K" & Summary_Table_Row).Value = percent_chn
      
      
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      
      
      
      
      
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset
      total_vol = 0
    ' If the cell immediately following a row is the same...
    
    
    Else

      ' Add to
      total_vol = total_vol + ws.Cells(i, 7).Value

    End If

   
Next i
Next ws


















End Sub