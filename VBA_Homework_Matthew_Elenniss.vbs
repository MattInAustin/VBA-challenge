Sub ticker():
''This VBA Scribt auto formats and assigns cells for the Multiple_Year_Stock_Data Excel Sheet''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Variables''
Dim tickerName As String
Dim changeRow As Integer
changeRow = 2
Dim tickerRow As Integer
tickerRow = 2
Dim openValue As Double
openValue = 0
Dim yearlyChange As Double
Dim openPrice As Double
Dim closePrice As Double
Dim percentChange As Double
Dim percentRow As Integer
percentRow = 2
Dim totalVolume As Double
totalVolume = 0
Dim volumeRow As Integer
volumeRow = 2
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Apply ticker names''
    For i = 2 To 44000
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            tickerName = Cells(i, 1).Value
            Cells(tickerRow, 9).Value = tickerName
            tickerRow = tickerRow + 1
        End If
    Next i
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Calculate Year Change''
    For i = 2 To 44000
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            openPrice = Cells(i, 3).Value
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            closePrice = Cells(i, 6).Value
            yearlyChange = openPrice - closePrice
            Cells(changeRow, 10).Value = yearlyChange
            changeRow = changeRow + 1
        End If
    Next i
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Calculate Percent Change''
    For i = 2 To 44000
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            openPrice = Cells(i, 3).Value
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            closePrice = Cells(i, 6).Value
            percentChange = (openPrice - closePrice) / closePrice
            'Worksheets.Cells(percentRow, 11).Style = "Percent"'
            Cells(percentRow, 11).Value = percentChange
            percentRow = percentRow + 1
        End If
    Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Calc total volume.''
For i = 2 To 44000
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            totalVolume = totalVolume + Cells(i, 7).Value
            Cells(volumeRow, 12).Value = totalVolume
            volumeRow = volumeRow + 1
            totalVolume = 0
        Else
            totalVolume = totalVolume + Cells(i, 7).Value
        End If
    Next i
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub
''The End''
''Code by Matthew Elenniss''
''5/14/2020''