Attribute VB_Name = "Module1"
Sub ticker_kicker()

  ' Est variables for your nouns like ticker, the eventual total value of each ticker,
  ' and the row of the table we're building to hold the answers so we can direct our output to the right spot
  ' Est type here, too e.g. 'as string' 'as double' 'as long' w/ each variable so you may not have to later.
  ' Hello, computer.
  ' Brady sets variable value proximate to declaration. Why did i have to declare ws when in the Wells example we did not?
  ' And for that matter 'bottom' which replaced LastRow in the Wells example...but was never defined as ... anything?
  

Dim TICKER As String
Dim TxTOTAL As Double
Dim TCKRtableROW As Integer
Dim BOTTOM As Long
' Dim ws As Worksheets

TxTOTAL = 0
TCKRtableROW = 2
BOTTOM = Cells(Rows.Count, 1).End(xlUp).Row

' Loop through the data, pick up volumes for each ticker, totaling along the way.
' As each ticker's final total is determined, send its symbol and final total to the to-be-built table
' Be sure to reset the target for the next output & clear your calculator w/ each new ticker symbol
' create your table's column headers first

Cells(1, 11).Value = ("TICKER")
Cells(1, 12).Value = ("TOTAL VOLUME")

For i = 2 To BOTTOM
  
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      TICKER = Cells(i, 1).Value
      TxTOTAL = TxTOTAL + Cells(i, 7).Value
      
      Range("K" & TCKRtableROW).Value = TICKER
      Range("L" & TCKRtableROW).Value = TxTOTAL
      
      TCKRtableROW = TCKRtableROW + 1
      TxTOTAL = 0

    Else

      TxTOTAL = TxTOTAL + Cells(i, 7).Value

    End If

  Next i

End Sub
