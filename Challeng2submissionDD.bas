Attribute VB_Name = "Module1"

Sub Challenge2submissionDD()


Dim Wrksht As Worksheet
Dim ticker As String
Dim OpngPrce As Double 'double cause it contains large date
Dim ClsngPrce As Double
Dim YrlyChnge As Double
Dim prcntchnge As Double
Dim volume As Double
Dim i As Double
Dim lastrow As Double
Dim sumtablerow As Double

'Finally checking all sheets

For Each Wrksht In Worksheets
   Wrksht.Cells(1, 9) = "Ticker"
   Wrksht.Range("I1").ColumnWidth = 10
   Wrksht.Cells(1, 10) = "Yearly Change"
   Wrksht.Range("J1").ColumnWidth = 14
   Wrksht.Cells(1, 11) = "Percentage Change"
   Wrksht.Range("K1").ColumnWidth = 18   'initialise all header inputs and column width
   Wrksht.Cells(1, 12) = "Total Volume"
   Wrksht.Range("L1").ColumnWidth = 18
   Wrksht.Cells(1, 15) = "Ticker"
   Wrksht.Cells(1, 16) = "Value"


   Wrksht.Range("P1").ColumnWidth = 17
   Wrksht.Cells(2, 14) = "Greatest % Increase"
   Wrksht.Range("N1").ColumnWidth = 22
   Wrksht.Cells(3, 14) = "Greatest % Decrease"
   Wrksht.Cells(4, 14) = "Greatest Total Volume"


ticker = " "
sumtablerow = 2
OpngPrce = 0
ClsngPrce = 0
lastrow = Wrksht.Cells(Rows.Count, 1).End(xlUp).Row

 'This loop finds the values


For i = 2 To lastrow
OpngPrce = Wrksht.Cells(sumtablerow, 3).Value  'initiate variable opens with data

'This changes to the next value
 If Wrksht.Cells(i + 1, 1).Value <> Wrksht.Cells(i, 1).Value Then

 ticker = Wrksht.Cells(i, 1).Value
 volume = volume + Wrksht.Cells(i, 7).Value
 ClsngPrce = Wrksht.Cells(i, 6).Value
 YrlyChnge = ClsngPrce - OpngPrce
 prcntchnge = YrlyChnge / OpngPrce   '% changed

 Wrksht.Range("I" & sumtablerow).Value = ticker
 Wrksht.Range("J" & sumtablerow).Value = YrlyChnge
 Wrksht.Range("K" & sumtablerow).Value = prcntchnge      'to Print out the result data values
 Wrksht.Range("K" & sumtablerow).Style = "Percent"
 Wrksht.Range("L" & sumtablerow).Value = volume

 sumtablerow = sumtablerow + 1
 volume = 0

 Else
  volume = volume + Wrksht.Cells(i, 7).Value

 End If

Next i



Dim yearchangerow As Double
Dim r As Double
Dim maxIncrease As Double
Dim maxDecrease As Double
Dim maxVolume As Double

maxIncrease = 0
maxDecrease = 0
maxVolume = 0

yearchangerow = Wrksht.Cells(Rows.Count, 10).End(xlUp).Row 'selects up

' starts the loop thru the whole result data
For r = 2 To yearchangerow

' first IF statements to color format the data on the result data
If Wrksht.Cells(r, 10) >= 0 Then
Wrksht.Cells(r, 10).Interior.Color = RGB(109, 255, 109)

Else
Wrksht.Cells(r, 10).Interior.Color = RGB(239, 13, 13) 'when the value is below 0, negative

End If


If maxIncrease < Wrksht.Cells(r, 11).Value Then

maxIncrease = Wrksht.Cells(r, 11).Value

Wrksht.Cells(2, 15).Value = Wrksht.Cells(r, 9).Value
Wrksht.Cells(2, 16).Value = maxIncrease               'Prints results at another table
Wrksht.Cells(2, 16).Style = "Percent"

ElseIf maxDecrease > Wrksht.Cells(r, 11).Value Then

maxDecrease = Wrksht.Cells(r, 11).Value
Wrksht.Cells(3, 15).Value = Wrksht.Cells(r, 9).Value
Wrksht.Cells(3, 16).Value = maxDecrease               'Prints results at another table
Wrksht.Cells(3, 16).Style = "Percent"

End If

'Another conditional to get Ticker with the maxVolume,
'similar formula with maxIncrease as to find the biggest value in the result data
If maxVolume < Wrksht.Cells(r, 12).Value Then
maxVolume = Wrksht.Cells(r, 12).Value
Wrksht.Cells(4, 15).Value = Wrksht.Cells(r, 9).Value
Wrksht.Cells(4, 16).Value = maxVolume                 'Prints results at another table

End If


Next r


Next Wrksht



End Sub



