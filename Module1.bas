Attribute VB_Name = "Module1"
Sub Multi_year_data():

For Each ws In Worksheets

  total_stock_volume = 0

  Summary_Table_Row = 2
  
  Lrow = Cells(Rows.Count, 1).End(xlUp).Row
  
  Range("I1").Value = "Ticker"
  
  Range("L1").Value = "Total Stock Volume"
  
  Range("J1").Value = "yearly change"
  
  Range("K1").Value = "percent change"
  
  '-----------------------------------------------------
  
  'Dim Opening_Price As Double
  
  Dim current_stock As Boolean
  
  Dim ticker As String
  
  Dim total_stock_value, yearly_change, Percent_Change, Opening_Price As Long
  
  total_stock_volume = 0
       
  For i = 2 To Lrow
    
  ' find current stock
  
    If current_stock = False Then
         
         Opening_Price = Cells(i, 3).Value
         
         current_stock = True
    
   End If

 ' find when stock change
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

 
      ticker = Cells(i, 1).Value

   
      total_stock_volume = total_stock_volume + Cells(i, 7).Value

     
      Range("I" & Summary_Table_Row).Value = ticker

     
      Range("L" & Summary_Table_Row).Value = total_stock_volume

                 
            
      Opening_Price = Cells(i, 3).Value

    
      yearly_change = Cells(i, 6).Value - Opening_Price

     
      Range("J" & Summary_Table_Row).Value = yearly_change

      
      Percent_Change = (yearly_change / Opening_Price) * 100

      
      Range("K" & Summary_Table_Row).Value = Percent_Change

      
      Summary_Table_Row = Summary_Table_Row + 1
    
      
      stockPriceAlreadyCaptured = False
    
    Else

    
      total_stock_volume = total_stock_volume + Cells(i, 7).Value

    End If

  Next i
  
Next ws

End Sub

