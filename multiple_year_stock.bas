Attribute VB_Name = "Module1"
' Steps:

      '1. Create ticker symbol
      '2. Display yearly change from opening price at beginning of year to closing price at end of year
      '3. Display percent change from opening price of a given year to closing price at end of year
      '4. Display total stock volume of stock

Sub multiple_year_stock_data():

' assign variable to hold worksheet name
Dim ws As Worksheet

' loop through all worksheets in Workbook
For Each ws In Worksheets

' find last row (last row of numbers)
Dim LastRow As LongLong
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' define FirstRow
Dim FirstRow As Integer
FirstRow = 2

' Create Output Column Headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

' set variables needed for calculations
Dim ticker As String
ticker = " "
Dim openprice As Double
openprice = 0
Dim closeprice As Double
closeprice = 0
Dim pricechange As Double 'price change is yearly
pricechange = 0
Dim percentagechange As Double  'percentage change is yearly
percentagechange = 0
Dim totalvolume As Double
totalvolume = 0

' set location for variables
Dim Summarytable As Long
Summarytable = 2

' set location for openprice (initial stock value)
openprice = ws.Range("C2").Value

' loop through rows
For Row = FirstRow To LastRow

' Loop through tickers setting code to switch ticker when value changes
If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
    
'Set ticker
ticker = ws.Cells(Row, 1).Value
        
' Calculate closing price
closeprice = ws.Cells(Row, 6).Value

' Calculate yearly price change
pricechange = closeprice - openprice
        
' Ensure openprice is not equal to zero due to you dividing by this number
If openprice <> 0 Then
percentagechange = (pricechange / openprice) * 100
        
End If
         
'Add to ticker volume
totalvolume = totalvolume + ws.Cells(Row, 7).Value
        
' Print ticker in Column I
ws.Range("I" & Summarytable).Value = ticker
        
' Print price change in Column J
ws.Range("J" & Summarytable).Value = pricechange
        
' Set price to colorcode to red if it is a price decrease; set to green if it is a price increase
ws.Range("J" & Summarytable).Value = pricechange
If pricechange > 0 Then
ws.Range("J" & Summarytable).Interior.ColorIndex = 4
ElseIf pricechange <= 0 Then
ws.Range("J" & Summarytable).Interior.ColorIndex = 3
            
End If
        
' Print price change as a percent in column K
ws.Range("K" & Summarytable).Value = (CStr(percentagechange) & "%")
        
' Print total stock volume in column L
 ws.Range("L" & Summarytable).Value = totalvolume
        
' Add 1 to summary table count to begin on row 2
Summarytable = Summarytable + 1

' Calculate next open price
openprice = ws.Cells(Row + 1, 3).Value
        
End If

Next Row

Next ws

End Sub
