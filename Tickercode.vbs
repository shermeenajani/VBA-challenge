Option Explicit

Sub vbachallengehw()

'dim variables
'Set ticker name as a string
Dim Ticker As String
'Set OpeningPrice from the beginning of a year as double
Dim OpeningPrice As Double
'Set ClosingPrice from the end of the year as double
Dim ClosingPrice As Double
'Set percentchange from opening price to the closing price as a year as a double
Dim PercentChange As Double
'Set the yearlychange from the opening price to the closing price for the year as a double
Dim YearlyChange As Double
'Set total stock volume of the stock as double
Dim Volume As Double
'Set LastRow for the number of rows in the worksheet as double
Dim LastRow As Double
'Set i as an double to loop through the rows
Dim i As Double
'Set j as an double to loop through the worksheets
Dim j As Double
'Set the summarytablerow for the data in the table for the analysis
Dim SummaryTableRow1 As Integer
'Set the summarytablerow for the data in the table for the bonus
Dim SummaryTableRow2 As Integer
'Set the variable to hold the Greatest % Increase
Dim GreatestIncrease As Double
'Set the variable to hold the Greatest % Increase Ticker
Dim GreatestIncreaseTicker As String
'Set the variable to hold the Greatest % Decrease
Dim GreatestDecrease As Double
'Set the variable to hold the Greatest % Decrease Ticker
Dim GreatestDecreaseTicker As String
'Set the variable to hold the Greatest Volume
Dim GreatestVolume As Double
'Set the variable to hold the Greatest Volume Ticker
Dim GreatestVolumeTicker As String

'Loop through the worksheets
For j = 1 To Sheets.Count

'Activate the current worksheet
Worksheets(Sheets(j).Name).Activate

'Find the number of rows in the worksheet
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Create headers for the analysis table in the worksheet
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'Create headers and columns for the bonus table in the worksheet
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

'clear all counters
Volume = 0

'Track each ticker symbol data in the summary table for the analysis
SummaryTableRow1 = 2

'Track each ticker symbol data in the summary table for the bonus
SummaryTableRow2 = 2

'Intiatze the counters for the Greatest % Increase, Greatest % Decrease and Greatest Volume for the year
GreatestIncrease = 0
GreatestDecrease = 0
GreatestVolume = 0

' Loop through the rows in the worksheet
For i = 2 To LastRow

'If the ticker symbol is not the same as the previous cell then save the opening price
If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then

'Save the OpeningPrice
OpeningPrice = Cells(i, 3).Value

End If


'If the ticker symbol is not the same as the next cell then save the ticker symbol, opening price and volume
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'Save the Ticker Symbol
Ticker = Cells(i, 1).Value

'Add to the volume
Volume = Volume + Cells(i, 7).Value

'add the closingprice
ClosingPrice = Cells(i, 6).Value

'calculate yearlychange
YearlyChange = ClosingPrice - OpeningPrice

'calculate percentchange
'Calculation for the PercentChange
'If OpeningPrice = 0 then PercentChange = 0
If OpeningPrice = 0 Then
PercentChange = 0
    Else
    
    'Otherwise calculate the PercentChange based on the equation
    PercentChange = (ClosingPrice - OpeningPrice) / (OpeningPrice)
    
End If


'conditional formating for yearly change in price (green if positive, red if negative)
'Color formatting for the PercentChange
'If the PercentChange is positive then the cell will be green
If PercentChange >= 0 Then
Range("K" & SummaryTableRow1).Interior.ColorIndex = 4
Else

'If the PercentChange is negative then the cell will be red
Range("K" & SummaryTableRow1).Interior.ColorIndex = 3

End If

' Print the Ticker in the Summary Table
Range("I" & SummaryTableRow1).Value = Ticker

' Print the Yearly Change in the Summary Table
Range("J" & SummaryTableRow1).Value = YearlyChange
      
' Print the PercentChange in the Summary Table
Range("K" & SummaryTableRow1).Value = PercentChange

'Calculate the GreatestIncrease
'If PercentChange is greater than the Greatest % Increase value then
If PercentChange > GreatestIncrease Then

'replace the Greatest % Increase value with the PercentChange value
GreatestIncrease = PercentChange
'and replace the GreatestIncreaseTicker with the Ticker value
GreatestIncreaseTicker = Ticker

End If


'Calculate the GreatestDecrease
'If PercentChange is less than than the Greatest % Decrease value then
If PercentChange < GreatestDecrease Then

'replace the Greatest % Decrease value value with the PercentChange value
GreatestDecrease = PercentChange
'and replace the GreatestDecreaseTicker with the Ticker value
GreatestDecreaseTicker = Ticker

End If
      
' Print the Volume in the Summary Table
Range("L" & SummaryTableRow1).Value = Volume

'Calculate the GreatestVolume
'If Volume is greater than than the Greatest Volume value then
If Volume > GreatestVolume Then

'replace the Greatest Volumen value value with the Volume value
GreatestVolume = Volume
'and replace the GreatestVolumeTicker with the Ticker value
GreatestVolumeTicker = Ticker

End If

' Add one to the summary table row for the analysis
SummaryTableRow1 = SummaryTableRow1 + 1

      
' Reset the Volume
Volume = 0

    'If the ticker symbol is the same as the previous cell then save the volume
    Else
      ' Add to the Volume Total
      Volume = Volume + Cells(i, 7).Value

End If

'Go to the next row
Next i

'Format the column to use percentages
Range("K2:K" & LastRow).NumberFormat = "0.00%"

'Print the Greatest Increase Ticker
Cells(2, 16).Value = GreatestIncreaseTicker
'Print Greatest Decrease Ticker
Cells(3, 16).Value = GreatestDecreaseTicker
'Print Greatest Volume Ticker
Cells(4, 16).Value = GreatestVolumeTicker
'Print Greatest % Increase
Cells(2, 17).Value = GreatestIncrease
'Format the Greatest % Increase cell to display %
Cells(2, 17).NumberFormat = "0.00%"
'Print Greatest % Decrease
Cells(3, 17).Value = GreatestDecrease
'Format the Greatest % Decrease cell to display %
Cells(3, 17).NumberFormat = "0.00%"
'Print Greatest Volume
Cells(4, 17).Value = GreatestVolume
'Format the Greatest % Volume cell to display Scientific Method
Cells(4, 17).NumberFormat = "0.0000E+00"

'Make sure the vlaues in columns I through Q fit in the columns
Worksheets(Sheets(j).Name).Columns("I:Q").AutoFit

'Go to the next worksheet
Next j

End Sub

