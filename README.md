# VBA-challenge
Sub Stock_Data()

Dim Stock_Data As String
For Each ws In Worksheets

'Set variables
Dim Summary_Table_Row As Integer
Dim Ticker As String
Dim Total_Stock_Volume As LongLong
Dim yearlyChange As Double
Dim percentChange As Double
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim Max As Double
Dim Min As Double

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Summary_Table_Row = 2


'Loop through the Ticker and Stock volume
For i = 2 To LastRow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

Ticker = ws.Cells(i, 1).Value

ws.Cells(i, 9).Value = Ticker

ws.Range("I1").Value = "Ticker"

ws.Range("I" & Summary_Table_Row).Value = Ticker

Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

ws.Range("L1").Value = "Total Stock Volume"

ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

Summary_Table_Row = Summary_Table_Row + 1

Total_Stock_Volume = 0

Else
Total_Stock_Volume = (Total_Stock_Volume + ws.Cells(i, 7).Value)

End If
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ClosePrice = ws.Cells(i, 6).Value
OpenPrice = ws.Cells(i, 3).Value
yearlyChange = ClosePrice - OpenPrice
ElseIf OpenPrice <> 0 Then
percentChange = (yearlyChange / OpenPrice) * 100


ws.Range("J1").Value = "Yearly Change"
ws.Range("J" & Summary_Table_Row).Value = yearlyChange
ws.Range("K1").Value = "Percent Change"
ws.Range("K" & Summary_Table_Row).Value = percentChange

End If

Next i


'Set conditional formatting
BottomRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To BottomRow

If ws.Cells(i, 10).Value >= 0 Then

ws.Cells(i, 10).Interior.ColorIndex = 4

Else
ws.Cells(i, 10).Interior.ColorIndex = 3

End If
Next i

Next ws

End Sub



