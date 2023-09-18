Attribute VB_Name = "Module1"
Sub Stock_Analysis()
' Set an initial variable for holding the brand name
  Dim Ticker_name As String
  Dim ws As Worksheet
  Dim RowCount As Long
  ' Set an initial variable for holding the total per credit card brand
   Dim Stock_Total As Double
   Stock_Total = 0
  Dim opening_price As Double
   Dim closing_price As Double

For Each ws In ActiveWorkbook.Worksheets

'Set ws = Sheet1
ws.Cells(1, 10).Value = ("Ticker")
ws.Cells(1, 11).Value = ("Yearly_Change")
ws.Cells(1, 12).Value = ("Percentage_Change")
ws.Cells(1, 13).Value = ("Total_Stock_Volume")
ws.Cells(2, 14).Value = ("Greatest %increase ")
ws.Cells(3, 14).Value = ("Greatest % decline")
ws.Cells(4, 14).Value = ("Greatest stock volume")

ws.Cells(1, 15).Value = ("Ticker")


  RowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).EntireRow.Row
    
      ' Keep track of the location for each credit card brand in the summary table
      Dim Summary_Table_Row As Integer
      Summary_Table_Row = 2
    
      ' Loop through all stocks for one year
      
      For i = 2 To RowCount
    
        ' Check if we are still within the ticker, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
          ' Set the Ticker name
          Ticker_name = ws.Cells(i, 1).Value
    
          ' Add to the Stock_Total
          Stock_Total = Stock_Total + ws.Cells(i, 7).Value
    
          ' Print the Ticker in the Summary Table
          ws.Range("J" & Summary_Table_Row).Value = Ticker_name
    
          ' Print the Stock_Total Amount to the Summary Table
          ws.Range("M" & Summary_Table_Row).Value = Stock_Total
    
          ' Add one to the summary table row
    
    
    
         opening_price = ws.Cells(i, 3).Value

       closing_price = ws.Cells(i, 6).Value
    
      ws.Range("K" & Summary_Table_Row).Value = closing_price - opening_price
    
      ws.Range("L" & Summary_Table_Row).Value = ((Range("K" & Summary_Table_Row).Value / opening_price) * 100)
    Summary_Table_Row = Summary_Table_Row + 1
    
      End If
    
      If ws.Cells(i, 11).Value < 0 Then
        
        ws.Cells(i, 11).Interior.ColorIndex = 3
    
      Else
       
       ws.Cells(i, 11).Interior.ColorIndex = 4
    
      End If
  
 Next i




ws.Cells(2, 16).Value = Application.WorksheetFunction.Max(ws.Columns("L"))

ws.Cells(3, 16).Value = Application.WorksheetFunction.Min(ws.Columns("L"))

ws.Cells(4, 16).Value = Application.WorksheetFunction.Max(ws.Columns("M"))
'    increase_number = WorksheetFunction.Match(ws.Cells(2, 16).Value)
'    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Columns("L"))
'
'    volume_number = WorksheetFunction.Match(WorksheetFunction.Max((ws.Columns("G"))
' final ticker symbol for  total, greatest % of increase and decrease, and average
'    ws.Cells(2, 15).Value = WorksheetFunction.Match(ws.Cells(2, 16).Value, ws.Cells(i, 10).Value, 0)
'    ws.Cells(3, 15).Value = WorksheetFunction.Match(ws.Cells(3, 16).Value, ws.Range("K"), 0)
'    ws.Cells(4, 15).Value = WorksheetFunction.Match(ws.Cells(3, 16).Value, ws.Range("L"), 0)

'ws.Cells(2, 15).Value = WorksheetFunction.Match(ws.Cells(2, 16).Value, ws.Range("J:J"), 0)
'ws.Cells(3, 15).Value = WorksheetFunction.Match(ws.Cells(3, 16).Value, ws.Range("K:K"), 0)
'ws.Cells(4, 15).Value = WorksheetFunction.Match(ws.Cells(3, 16).Value, ws.Range("L:L"), 0)





  
Next ws

End Sub


