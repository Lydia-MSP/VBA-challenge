Attribute VB_Name = "Module1"

'Instructions

'* Create a script that will loop through all the stocks for one year for each run and take the following information. Done!

  '* The ticker symbol. Done!

 '* Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  '* The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  '* The total stock volume of the stock. Done!

'* You should also have conditional formatting that will highlight positive change in green and negative change in red.
' "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".





Sub StockChallenge()

  ' Set an initial variable for holding the ticker name
  Dim Ticker_Name As String

  ' Set an initial variable for holding the total per tiker
  Dim Ticker_Total As Double
  Ticker_Total = 0
  Dim LastRow As Long
  Dim ws As Worksheet
  Dim starting_position As Long
  Dim Ticker_closing_price As Double
  Dim Ticker_opening_price As Double
  Dim Yearly_Change As Double
  Dim percentage_change As Double
  Dim Greatest_percentage_increase As Double
  Dim Greatest_percentage_decrease As Double
  Dim Greatest_total_value As Double
  
    

  ' Keep track of the location for each tiker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
 
'For Each ws In Worksheets
' Determine the Last Row

LastRow = Cells(Rows.Count, 1).End(xlUp).Row
starting_position = 2
  ' Loop through all ticker changes
  Ticker_opening_price = Cells(starting_position, 3)
  Range("J1").Value = "Volume"
  Range("I1").Value = "Ticker"
  Range("K1").Value = "Yearly Change"
  Range("L1").Value = "Percentage Change"
  
  
  
  For i = 2 To LastRow
    

    ' Check if we are still within the same tickert name, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the ticker name
      Ticker_Name = Cells(i, 1).Value
      starting_position = i + 1
      'Set closing price
      Ticker_closing_price = Cells(i, 6)
     
      
      'calculate yearly chage
    
      Yearly_Change = Ticker_closing_price - Ticker_opening_price
      If Ticker_opening_price = 0 Then
        percentage_change = 0
      Else
      
      percentage_change = Yearly_Change / Ticker_opening_price
      End If

       'Set opening_price
      Ticker_opening_price = Cells(starting_position, 3)
      ' Add to the ticker Total
      Ticker_Total = Ticker_Total + Cells(i, 7).Value

      ' Print the Ticker name in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Ticker to the Summary Table & yearly change & percentage
      Range("j" & Summary_Table_Row).Value = Ticker_Total
      Range("K" & Summary_Table_Row).Value = Yearly_Change
      If Yearly_Change < 0 Then
      Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
      Else
      Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
      End If
      
      Range("L" & Summary_Table_Row).Value = percentage_change
      

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
        
      ' Reset the Brand Total
      Ticker_Total = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      Ticker_Total = Ticker_Total + Cells(i, 7).Value
      
         
    
    End If
    
  Next i
  Greatest_percentage_increase = Cells(2, 11).Value
  Greatest_percentage_decrease = Cells(2, 11).Value
  Greatest_total_value = Cells(2, 10).Value
  
 LastRow = Cells(Rows.Count, 9).End(xlUp).Row
 For i = 2 To LastRow
 If Cells(i, 11).Value > Greatest_percentage_increase Then
 Greatest_percentage_increase = Cells(i, 11).Value
 End If
 If Cells(i, 11).Value < Greatest_percentage_decrease Then
 Greatest_percentage_decrease = Cells(i, 11).Value
 End If
 If Cells(i, 10).Value > Greatest_total_value Then
 Greatest_total_value = Cells(i, 10).Value
 End If
 Next i
 
 Cells(2, 14).Value = Greatest_percentage_increase
 Cells(3, 14).Value = Greatest_percentage_decrease
 Cells(4, 14).Value = Greatest_total_value
 Cells(2, 13).Value = "Greatest_percentage_increase"
 Cells(3, 13).Value = "Greatest_percentage_decrease"
 Cells(4, 13).Value = "Greatest_total_value"

 

End Sub

