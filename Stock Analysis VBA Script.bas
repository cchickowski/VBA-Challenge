Attribute VB_Name = "Module1"
'Create a script that will loop through all the stocks for one year and output the following information
'The ticker symbol.
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.

Sub Multi_Year_Stock_data()


' Set an initial variable for holding the ticker name
  Dim ticker As String
  ticker = " "
  
  'Set worksheet variable
  Dim ws As Worksheet
  

  ' Set an initial variable for total volume of the stocks
  Dim Volume As Double
  Volume = 0

  
  'set variables for yearly change and percent change
  Dim year_open As Double
  year_open = 0
  
  Dim year_close As Double
  year_close = 0
  
  Dim yearly_change As Double
  yearly_change = 0
  
  Dim percent_change As Double
  percent_change = 0
  
  
  'Overflow Error fix
  On Error Resume Next
  
  
  
  'variable for summary table
  Dim Summary_Table_Row As Integer
  
  
'worskheet loop
 For Each ws In ThisWorkbook.Worksheets
  'column headers
  ws.Cells(1, 11).Value = "ticker"
  ws.Cells(1, 12).Value = "Stock Volume"
  ws.Cells(1, 13).Value = "Yearly Change"
  ws.Cells(1, 14).Value = "Percent Change"
  
  
  Summary_Table_Row = 2
  
  'Set initial value for open price first stock on WS
  year_open = ws.Cells(2, 3).Value
  
  
  ' Loop through all stocks on worksheets
  
    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    


        ' Check if we are still within the same stock ticker, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        ' Set the ticker start point
        ticker = ws.Cells(i, 1).Value
      
        'Determine year_change
        year_close = ws.Cells(i, 6).Value
      
        yearly_change = year_close - year_open
      
        percent_change = ((year_close - year_open) / year_open)
      
        
      
      
        ' Add to the ticker Total
        Volume = Volume + ws.Cells(i, 7).Value
      
      

        ' Print the ticker in the Summary Table
        ws.Range("K" & Summary_Table_Row).Value = ticker

        ' Print the Volume to the Summary Table
        ws.Range("L" & Summary_Table_Row).Value = Volume
      
        'Print the yearly change in Summary Table and color code
        ws.Range("M" & Summary_Table_Row).Value = yearly_change
            If (yearly_change > 0) Then
            ws.Range("M" & Summary_Table_Row).Interior.ColorIndex = 4
            ElseIf (yearly_change <= 0) Then
            ws.Range("M" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
        
      
        'Print the percent change in Summary Table
        ws.Range("N" & Summary_Table_Row).Value = percent_change
        ws.Range("N" & Summary_Table_Row).NumberFormat = "0.00%"
      
      
        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
      
        'Get the next open price
        year_open = ws.Cells(i + 1, 3).Value
      
        'reset values
        percent_change = 0
        Volume = 0
      
      
      
      

        ' If the cell immediately following a row is the same ticker...
        Else
        

        ' Add to the ticker Total
        Volume = Volume + ws.Cells(i, 7).Value
    
    
    
    
        End If
        
    

    Next i
    
 Next ws
 


End Sub

