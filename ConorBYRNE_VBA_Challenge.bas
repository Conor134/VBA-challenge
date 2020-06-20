Attribute VB_Name = "Module1"
Option Explicit
Sub AllSheets()

Dim WS_Count As Integer
Dim i As Integer

' Get number of Worksheets
WS_Count = ThisWorkbook.Worksheets.Count

'Loop through the worksheets, make active and run various Subs
For i = 1 To WS_Count

    ThisWorkbook.Worksheets(i).Activate

'Calling Subs to run on each Sheet
     Call GetUnique
     Call Formating
     Call MaxMinColourGreatest
     Call InsertText
     Call AutoFitColumns
     Call DeleteRows
     

Next i


End Sub

Sub GetUnique()

' Set  variables
Dim StockName As String
Dim i, StartYearPrice, EndYearPrice, Price As Long
Dim DatePrices, DeltaChange, PercentageChange, TotVol, ChangingRow, Diff, LastRow As Double

'Set Total Volume as 0 before running the loop
 TotVol = 0
 
' Keeping track of the location in the main table and the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
ChangingRow = 1
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

' Loop through all the dailty Stock prices
  For i = 2 To LastRow
   
' Check if we are still within the same stock, if we are not...

    
      If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
 
      'Calculating the Row vlues of the first and last day for "i" Stock
      Diff = i - ChangingRow
      
      ChangingRow = i - Diff + 1
      
     'Getting Stockname, Starting and Closing Prices
      StartYearPrice = Cells(ChangingRow, 3).Value
      StockName = Cells(i, 1).Value
      EndYearPrice = Cells(i, 6).Value

      'Calaulating Total Volume
      TotVol = TotVol + Cells(i, 7).Value

      ' Print the StockName in the Summary Table
      Range("J" & Summary_Table_Row).Value = StockName
      
      'Calculating the DeltaChange between start and end of the year
       DeltaChange = EndYearPrice - StartYearPrice
       
      ' Print Starting and Closing Prices and DeltaChange in Summary Table
      Range("K" & Summary_Table_Row).Value = StartYearPrice
      Range("L" & Summary_Table_Row).Value = EndYearPrice
      Range("M" & Summary_Table_Row).Value = DeltaChange
      'Skipping the problem of having a 0 price for End Year Price
      If StartYearPrice = 0 Then
        Range("N" & Summary_Table_Row).Value = 0
      Else
         Range("N" & Summary_Table_Row).Value = DeltaChange / StartYearPrice
      End If
      ' Print Total Volume in Summary Table
      Range("O" & Summary_Table_Row).Value = TotVol
    
    ' Add one to the summary table row and restting ChangingRow Table
      Summary_Table_Row = Summary_Table_Row + 1
      ChangingRow = i
      
      ' Reset the Total Volume
      TotVol = 0
      
    ' If the cell immediately following a row is the same stock, continue
    Else
     
      ' Add to the Total Volume
      TotVol = TotVol + Cells(i, 7).Value
 
    End If

  Next i

End Sub

Sub InsertText()
'Inserting Title text to Summary Table
Cells(1, 10).Value = "StockCode"
Cells(1, 13).Value = "Delta Yearly Change"
Cells(1, 14).Value = "% Yearly Change"
Cells(1, 15).Value = "Total Stock Volume"
Cells(1, 11).Value = "Start of Year Price"
Cells(1, 12).Value = "End of Year Price"

'Inserting Title text to Greatest Table
Cells(1, 19).Value = "Ticker"
Cells(1, 20).Value = "Value"
Cells(2, 18).Value = "Greatest % Increase"
Cells(3, 18).Value = "Greatest % Decrease"
Cells(4, 18).Value = "Greatest Total Volume"


End Sub
Sub AutoFitColumn()

'Autofitting columns to fit the datalength
Columns("C:D").Select
Selection.ColumnWidth = 40


End Sub
Sub MaxMinColourGreatest()

Dim MaxIncrease, MinDecrease, x, MaxVol  As Double
Dim StockIncrease, StockDecrease, StockVol As String

'Setting values to prior to starting loop
        MaxIncrease = Cells(2, 14).Value
        MinDecrease = Cells(2, 14).Value
        MaxVol = Cells(2, 15).Value
        
' Starting Do loop until table runs out of stocknames
    x = 2
        
        Do Until IsEmpty(Cells(x, 10))
        
            'To see if Increase is > then set it
            If Cells(x, 14).Value > MaxIncrease Then
                 MaxIncrease = Cells(x, 14).Value
                 StockIncrease = Cells(x, 10).Value
                      
            End If
            
            'To see if Increase is < then set it
            If Cells(x, 14).Value < MinDecrease Then
                 MinDecrease = Cells(x, 14).Value
                 StockDecrease = Cells(x, 10).Value
            End If
            
            'To see if Volume is > then set it
            If Cells(x, 15).Value > MaxVol Then
                 MaxVol = Cells(x, 15).Value
                 StockVol = Cells(x, 10).Value
            End If
            
            'To see if %CHange is <0 then set cell colour to red
            If Cells(x, 14).Value < 0 Then
                Cells(x, 14).Select
                
              Cells(x, 14).Interior.ColorIndex = 3
              End If
              
              'To see if %CHange is >0 then set cell colour to green
             If Cells(x, 14).Value > 0 Then
                Cells(x, 14).Select
                
              Cells(x, 14).Interior.ColorIndex = 4
              End If
             x = x + 1
    
       Loop
        
'Putting the Greated vlaues into a table
Range("S" & 2).Value = StockIncrease
Range("T" & 2).Value = MaxIncrease

Range("S" & 3).Value = StockDecrease
Range("T" & 3).Value = MinDecrease
        
Range("S" & 4).Value = StockVol
Range("T" & 4).Value = MaxVol
 
End Sub
Sub Formating()

Columns("N:N").Select
Selection.Style = "Percent"
Selection.NumberFormat = "0.00%"

Range("T2:T3").Select
Selection.Style = "Percent"
Selection.NumberFormat = "0.00%"

Range("T4:T4").Select
Selection.NumberFormat = "0.00E+00"

End Sub
Sub AutoFitColumns()

    Cells.Select
    Cells.EntireColumn.AutoFit
End Sub
Sub DeleteRows()

    Columns("K:L").Select
    Selection.Delete Shift:=xlToLeft
End Sub
