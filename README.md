# VBA-Challenge
Sub Stock()

Dim Ticker As String
Dim Volume As Double
Dim PercentChange As Double
Dim YearlyChange As Double
Dim NextVolume As Double

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Volume = 0

For i = 2 To lastrow

  If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
    
    Ticker = Cells(i, 1).Value
    
     Cells(i, 10) = Ticker
       
End If
    
Next i

End Sub
