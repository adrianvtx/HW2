Attribute VB_Name = "Module1"
Sub Ticker() 'HW02-VBA-EASY

  ' Set initial variable for holding the Ticker
Dim TKR As String 'Ticker

Dim TTL As Double 'Total
TTL = 0

  ' Keep track of the row in the summary table
Dim STR As Integer 'STR- summary table row
STR = 2 'Begin on row 2

  ' Find last row on sheet
lastRow = Cells(Rows.Count, "A").End(xlUp).Row - 1

  ' Loop through all rows
  For i = 2 To lastRow

    ' Check if we are still within the same Ticker string value
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Places Ticker string value in TKR
      TKR = Cells(i, 1).Value

      ' Sums Volume into TKR
      TTL = TTL + Cells(i, 7).Value

      ' Print the Ticker string into Summary Table column I
      Range("I" & STR).Value = TKR

      ' Print the Total Volume into Summary Table column J
      Range("J" & STR).Value = TTL

      ' Increment STR- summary table row
      STR = STR + 1
      
      ' Reset Volume Variable to zero
      TTL = 0

    Else

      ' Add current volume to sum of Ticker volume
      TTL = TTL + Cells(i, 7).Value

    End If

  Next i

End Sub
