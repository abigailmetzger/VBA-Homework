Sub MultiYearData()
'Loop through all worksheets

For Each ws In Worksheets

'Define Variables
Dim Ticker As String
Dim Volume As Double

Dim Total_Stock_Volume As Double
Total_Table = 2

'Lable new headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Total Stock Volume"


lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through all ticker values and add up the volumes
For i = 2 To lastrow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Ticker = ws.Cells(i, 1).Value
    Volume = Volume + ws.Cells(i, 7).Value
    
    ws.Range("I" & Total_Table).Value = Ticker
    ws.Range("J" & Total_Table).Value = Volume
    
    Total_Table = Total_Table + 1
    Volume = 0

Else
    
    Volume = Volume + ws.Cells(i, 7).Value
    

End If

Next i

Next ws

End Sub
