Attribute VB_Name = "Module1"
Option Explicit
Sub launcher()
Formulario.Show
End Sub

Sub stocks_yearly()

'Varaibles declaration
Dim i, n As Long
Dim k, index As Integer
Dim tot_vol, open_price, pdone As Double
Dim ticker As String

'Cleaning old results
Columns("I:Q").Select
Selection.ClearContents
Selection.ClearFormats

'Initial settings
n = Cells(Rows.Count, 1).End(xlUp).Row
ticker = Range("A2")
open_price = Range("C2")
k = 2
tot_vol = 0

For i = 2 To n
pdone = i / n

    With Formulario
        .Label1.Caption = "Processing Row " & i & " of " & n
    End With

    If Cells(i, 1) = ticker Then
        tot_vol = tot_vol + Cells(i, 7).Value
    Else
        'Print result for the current ticker
        Cells(k, 9).Value = ticker
        Cells(k, 10).Value = Cells(i - 1, 6) - open_price
        If open_price = 0 Then
            Cells(k, 11) = 0
        Else
            Cells(k, 11).Value = Cells(k, 10) / open_price
        End If
        Cells(k, 12).Value = tot_vol
        
        'To start a new ticker
        ticker = Cells(i, 1).Value
        open_price = Cells(i, 3).Value
        k = k + 1
        tot_vol = Cells(i, 7).Value
    End If
    
    Cells(k, 9).Value = ticker
    Cells(k, 10).Value = Cells(i, 6) - open_price
    If open_price = 0 Then
        Cells(k, 11) = 0
    Else
        Cells(k, 11).Value = Cells(k, 10) / open_price
    End If
    Cells(k, 12).Value = tot_vol

Next i

Cells(2, 17) = Application.Max(Range(Cells(2, 11), Cells(k, 11)))
Cells(3, 17) = Application.Min(Range(Cells(2, 11), Cells(k, 11)))
Cells(4, 17) = Application.Max(Range(Cells(2, 12), Cells(k, 12)))

index = Application.Match(Cells(2, 17), Range(Cells(2, 11), Cells(k, 11)), 0) + 1
Cells(2, 16) = Cells(index, 9)

index = Application.Match(Cells(3, 17), Range(Cells(2, 11), Cells(k, 11)), 0) + 1
Cells(3, 16) = Cells(index, 9)

index = Application.Match(Cells(4, 17), Range(Cells(2, 12), Cells(k, 12)), 0) + 1
Cells(4, 16) = Cells(index, 9)

Call formats(k)

End Sub

Sub formats(nRows)

Dim i As Integer
' Formatting result  table
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

Cells(1, 16) = "Ticker"
Cells(1, 17) = "Value"
Cells(2, 15) = "Greatest % Increase"
Cells(3, 15) = "Greatest % Decrease"
Cells(4, 15) = "Greatest Total Volume"

For i = 2 To nRows

If Cells(i, 10) >= 0 Then
    Cells(i, 10).Interior.Color = RGB(0, 255, 0)
Else
    Cells(i, 10).Interior.Color = RGB(255, 0, 0)
End If

Cells(i, 11).NumberFormat = "0.0%"
Cells(2, 17).NumberFormat = "0.0%"
Cells(3, 17).NumberFormat = "0.0%"

Next i

End Sub

