# VBA-Challenge
Sub Stock()

For Each ws In Worksheets

'Summary Table Header
'Cells(1, 9).Value = "Ticker"
'Cells(1, 10).Value = "Yearly Change"
'Cells(1, 11).Value = "Percent Change"
'Cells(1, 12).Value = "Total Stock Volume"

'Label Cells in Each Worksheet
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"

'Label Cells Summary Table in Each Worksheet
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

Dim Ticker As String
Dim Volume As Double
Dim PercentChange As Double
Dim YearlyChange As Double
Dim NextVolume As Double
Dim Startrow As Double
Dim J As Double


lastRow = Cells(Rows.Count, 1).End(xlUp).Row

Volume = 0

Startrow = 2

J = 2


For i = 2 To lastRow

'Check to see if we are still in the same ticker

  If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
    Volume = Volume + Cells(i, 7)
    
        If Cells(Startrow, 3) = 0 Then
            For find_value = Startrow To i
            If Cells(find_value, 3).Value <> 0 Then
                            Startrow = find_value
            Exit For
            End If
            Next find_value
            End If
    
    YearlyChange = Cells(i, 6).Value - Cells(Startrow, 3).Value
    
    PercentChange = (YearlyChange / Cells(Startrow, 3).Value) * 100

    
    Cells(J, 10) = YearlyChange
    
    Cells(J, 11) = PercentChange
    
    Cells(J, 12) = Volume
    
    Cells(J, 9).Value = Cells(i, 1).Value

    Startrow = (i + 1)
    
    J = J + 1
    
    Volume = 0
    
    Else
    
    Volume = Volume + Cells(i, 7)
    
End If

If Cells(Startrow, 10).Value > 0 Then
            Cells(Startrow, 10).Interior.ColorIndex = 4
            
        ElseIf Cells(Startrow, 10).Value < 0 Then
            Cells(Startrow, 10).Interior.ColorIndex = 3
            
        Else
        
            Cells(Startrow, 10).Interior.ColorIndex = 6
            
        End If

Next i
            
Next ws

End Sub
