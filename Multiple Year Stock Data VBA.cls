VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Easy():

For Each ws In Worksheets
    With ws

Dim Ticker As String
Dim TotalStockVolume As Double
Dim StartRow As Integer

'Count the number of observations'
NumberObservation = .Cells(Rows.Count, 1).End(xlUp).Row

'Assign StartRow from 2'
StartRow = 2

'Assign the columns that contain Ticker and Total Stock Volume'
.Range("I1").Value = "Ticker"
.Range("J1").Value = "Total Stock Volume"

'Assign TotalStockVolume from 0'
TotalStockVolume = 0

'Make a Loop to create the column of Ticker and Total Stock Volume'
For i = 2 To NumberObservation:
    If .Cells(i, 1).Value = .Cells(i + 1, 1).Value Then
    TotalStockVolume = TotalStockVolume + .Cells(i + 1, 7).Value
    Else: Ticker = .Cells(i, 1).Value
            .Range("I" & StartRow) = Ticker
            TotalStockVolume = TotalStockVolume
            .Range("J" & StartRow) = TotalStockVolume
            StartRow = StartRow + 1
            TotalStockVolume = 0
    End If
Next i
End With
Next ws
End Sub

Sub Moderate1():

For Each ws In Worksheets
With ws

Dim OpenPrice As Double
Dim ClosePrice As Double

'Count the number of stocks'
NumberStockRow = .Cells(Rows.Count, 9).End(xlUp).Row

'Count the number of observations'
NumberObservation = .Cells(Rows.Count, 1).End(xlUp).Row

'Assign the columns that contain Yearly Change and Percent Change'
.Range("K1").Value = "Yearly Change"
.Range("L1").Value = "Percent Change"

'Assign the StartRow from 2'
StartRow = 2

'Make a loop to get the ClosePrice for each Stock'
For i = 2 To NumberObservation:
    If .Cells(i, 1).Value <> .Cells(i + 1, 1).Value Then
   ClosePrice = .Cells(i, 6).Value
    .Range("K" & StartRow) = ClosePrice
    StartRow = StartRow + 1
    End If
Next i
End With
Next ws
End Sub
Sub Moderate2():

For Each ws In Worksheets
With ws

Dim OpenPrice As Double
Dim ClosePrice As Double

'Count the number of stocks'
NumberStockRow = .Cells(Rows.Count, 9).End(xlUp).Row

'Count the number of observations'
NumberObservation = .Cells(Rows.Count, 1).End(xlUp).Row

'Assign the StartRow from 2'
StartRow = 2

'Make a loop to get the OpenPrice-ClosePrice and the Percent Change for each Stock'
For i = 2 To NumberObservation:
    If .Cells(i, 1).Value = .Range("I" & StartRow).Value Then
    OpenPrice = .Cells(i, 3).Value
    ClosePrice = .Range("K" & StartRow).Value
    .Range("K" & StartRow) = ClosePrice - OpenPrice
        If OpenPrice = 0 Then
        .Range("L" & StartRow) = 0
        Else
        .Range("L" & StartRow) = (ClosePrice - OpenPrice) / OpenPrice * 100 & "%"
        End If
    StartRow = StartRow + 1
    End If
Next i

'Make a loop to highlight to Yearly Change'
For i = 2 To NumberStockRow:
    If .Cells(i, 11).Value > 0 Then
    .Cells(i, 11).Interior.ColorIndex = 4
    ElseIf .Cells(i, 11).Value < 0 Then
    .Cells(i, 11).Interior.ColorIndex = 3
    End If
Next i
End With
Next ws

End Sub

Sub Hard():

For Each ws In Worksheets
With ws

'Assign the columns that contain Ticker and Value'
.Range("P1").Value = "Ticker"
.Range("Q1").Value = "Value"

'Assign the rows that contain Greatest % Increase , Greatest % Decrease and Greatest Total Volume'
.Range("O2").Value = "Greatest % Increase "
.Range("O3").Value = "Greatest % Decrease"
.Range("O4").Value = "Greatest Total Volume"

'Count the number of stocks'
NumberStockRow = .Cells(Rows.Count, 9).End(xlUp).Row

'Assign a maximum is 0'
Max = 0

'Make a loop to find the Greatest % Increase'
For i = 2 To NumberStockRow:
    If .Cells(i, 12).Value > Max Then
    Max = .Cells(i, 12).Value
    Stock = .Cells(i, 9).Value
    .Range("P2") = Stock
    .Range("Q2") = Max * 100 & "%"
    End If
Next i

'Assign a minimum is 0'
Min = 0

'Make a loop to find the Greatest % Decrease'
For i = 2 To NumberStockRow:
    If .Cells(i, 12).Value < Min Then
    Min = .Cells(i, 12).Value
    Stock = .Cells(i, 9).Value
    .Range("P3") = Stock
    .Range("Q3") = Min * 100 & "%"
    End If
Next i

'Assign a Max is 0'
Max = 0

'Make a loop to find the Greatest Total Volume'
For i = 2 To NumberStockRow:
    If .Cells(i, 10).Value > Max Then
    Max = .Cells(i, 10).Value
    Stock = .Cells(i, 9).Value
    .Range("P4") = Stock
    .Range("Q4") = Max
    End If
Next i

End With
Next ws
End Sub
