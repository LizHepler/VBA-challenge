Sub AgragateData()
Dim outputRow As Integer
Dim i, j As Long
Dim currentTicker As String
Dim tickers() As String
Dim firstRow As Long
Dim totalVol As LongLong

For Each quarter In Worksheets

    quarter.Activate

    i = 2
    outputRow = 2
    firstRow = 2
    totalVol = 0
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "Percent Change"
    Range("l1").Value = "Total Stock Volume"



    While (Not Cells(i, 1) = "")
        totalVol = totalVol + Cells(i, 7).Value
        If Not Cells(i, 1) = Cells(i + 1, 1) Then
            Cells(outputRow, 9).Value = Cells(i, 1).Value
            Cells(outputRow, 10).Value = Cells(i, 6).Value - Cells(firstRow, 3).Value
            Cells(outputRow, 11).Value = 1 - (Cells(firstRow, 3).Value / Cells(i, 6).Value)
            Cells(outputRow, 12).Value = totalVol
            outputRow = outputRow + 1
            firstRow = i + 1
            totalVol = 0
        End If
    
        i = i + 1
    Wend

    For Each cell In Range("J:J")
        If cell.Value > 0 Then
            cell.Interior.ColorIndex = 4
        ElseIf cell.Value < 0 Then
            cell.Interior.ColorIndex = 3
        End If
    Next cell

    Range("J1").Interior.ColorIndex = 0

    Range("J:J").NumberFormat = "0.00"
    Range("K:K").NumberFormat = "0.00%"

    j = 2
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"

    While (Not Cells(j, 9) = "")
        If Cells(j, 11).Value > Cells(2, 16).Value Then
            Cells(2, 15).Value = Cells(j, 9).Value
            Cells(2, 16).Value = Cells(j, 11).Value
        ElseIf Cells(j, 11).Value < Cells(3, 16).Value Then
            Cells(3, 15).Value = Cells(j, 9).Value
            Cells(3, 16).Value = Cells(j, 11).Value
        End If
        If Cells(j, 12).Value > Cells(4, 16).Value Then
            Cells(4, 15).Value = Cells(j, 9).Value
            Cells(4, 16).Value = Cells(j, 12).Value
        End If
        j = j + 1
    Wend

    Range("P2:P3").NumberFormat = "0.00%"
Next quarter
End Sub
