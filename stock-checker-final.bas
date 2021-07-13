Attribute VB_Name = "Module1"
Sub stock_checker_final()

'setting the column names
 Cells(1, 9).Value = "Ticker"
 Cells(1, 10).Value = "Yearly Change"
 Cells(1, 11).Value = "Percent Change"
 Cells(1, 12).Value = "Total Stock Volume"

  'Creating the variables
  Dim ticker_letter As String
  Dim init_amount As Double
  Dim fnl_amout As Double
  
  Dim year_amount As Double
  
  'Variable to hold percentage
  Dim year_cent As Long



  'Set initial variable for holding total volume per ticker
  Dim volume_total As Variant
  volume_total = 0

  'Keep track of location of each ticker in summary table
  Dim summary_table_row As Integer
  summary_table_row = 2
  Dim start_row As Long
  start_row = 2
  

  'Loop through all ticker types
  For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
  
  
  'Setting the initial and final amounts
        
   If Cells(i - 1, 1).Value <> Cells(i, 1).Value And Cells(i, 1).Value = Cells(i, 1).Value Then
            init_amount = Cells(i, 3).Value
   End If
        
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value And Cells(i, 1).Value = Cells(i, 1).Value Then
            fnl_amount = Cells(i, 6).Value
    End If
    
    year_amount = fnl_amount - init_amount
     
    'Fixing the the undefined issues
    If init_amount = 0 Then
      Cells(i, 11).Value = 1
      Else
     Cells(i, 11).Value = year_amount / init_amount
    End If
        

    'Check if we are still within the same ticket letter
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'Set the ticker letter
        ticker_letter = Cells(i, 1).Value
        
        'Add to the volume total
        volume_total = volume_total + Cells(i, 7).Value
        
        'start row
        start_row = i + 1
        
        'Print the ticker letter in the Summary Table
        Range("I" & summary_table_row).Value = ticker_letter
        
        'Print the volume total to the Summary Table
        Range("L" & summary_table_row).Value = volume_total
        
        'Print ticker letter and volume total
        Range("J" & summary_table_row).Value = year_amount
        
        'Print percent change to table
        Range("K" & summary_table_row).Value = "%" & Round((Cells(i, 11).Value * 100), 2)
        
        year_cent_row = year_cent_row + 1
        'Add one to the summary table
        summary_table_row = summary_table_row + 1
        
        'Reset the volume total
        volume_total = 0
        
     'If the cell immediately following a row is the same ticker
     Else
     
      'Add to the total volume
      volume_total = volume_total + Cells(i, 7).Value
        
     End If
    
  Next i
  
  'Adding the cell colors for negative and positive numbers
  For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
  
  If Cells(i, 10).Value >= 0 Then
                Cells(i, 10).Interior.Color = vbGreen
            Else
                Cells(i, 10).Interior.Color = vbRed
  End If
  Next i
    
End Sub


