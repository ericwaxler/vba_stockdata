Attribute VB_Name = "Module1"
Sub CalculateChange():
    Dim startPrice As Double
    Dim endPrice As Double
    Dim stockVolume As Single
    Dim tickerCounter As Integer
    Dim change As Double
    Dim pctChange As Double
     
    Dim bestTick As String
    Dim bestVal As Double
    Dim worstTick As String
    Dim worstVal As Double
    Dim volTick As String
    Dim volVal As Single
    
    Dim WS_Count As Integer
    Dim sht As Worksheet
    
    tickerCounter = 2
    endPrice = Cells(2, 6).Value
    WS_Count = ActiveWorkbook.Worksheets.Count

      
For h = 1 To WS_Count
   ThisWorkbook.Worksheets(h).Activate

    tickerCounter = 2
    startPrice = Cells(2, 3).Value

    Dim I As Long
        For I = 2 To Cells(Rows.Count, 1).End(xlUp).Row

            'check if ticker changes
            If Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
                endPrice = Cells(I, 6).Value
                change = endPrice - startPrice
                Cells(tickerCounter, 10).Value = change

                'conditionally format change
                If change >= 0 Then
                    Cells(tickerCounter, 10).Interior.ColorIndex = 4
                Else
                    Cells(tickerCounter, 10).Interior.ColorIndex = 3
                End If

                'insert stats
                
                
                pctChange = change / startPrice
                Cells(tickerCounter, 9).Value = Cells(I, 1).Value
                Cells(tickerCounter, 11).Value = Format(pctChange, "Percent")
                Cells(tickerCounter, 12).Value = stockVolume

                'check for extrema
                If pctChange > bestVal Then
                    bestVal = pctChange
                    bestTick = Cells(tickerCounter, 9).Value
                End If

                If pctChange < worstVal Then
                    worstVal = pctChange
                    worstTick = Cells(tickerCounter, 9).Value
                End If

                If stockVolume > volVal Then
                    volVal = stockVolume
                    volTick = Cells(tickerCounter, 9).Value
                End If

                'reset counters, update end price
                startPrice = Cells(I + 1, 3).Value
                stockVolume = 0
                tickerCounter = tickerCounter + 1
            Else
                If startPrice = 0 And Cells(I + 1, 3).Value <> 0 Then
                    startPrice = Cells(I + 1, 3).Value
                End If
                stockVolume = stockVolume + Cells(I, 7).Value
            End If
        Next I
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"


Cells(2, 16).Value = bestTick
Cells(2, 17).Value = Format(bestVal, "Percent")
Cells(3, 16).Value = worstTick
Cells(3, 17).Value = Format(worstVal, "Percent")
Cells(4, 16).Value = volTick
Cells(4, 17).Value = volVal

bestTick = 0
bestVal = 0
worstTick = 0
worstVal = 0
volTick = 0
volVal = 0
Next h

ThisWorkbook.Worksheets(1).Activate

'Cells(1, 16).Value = "Ticker"
'Cells(1, 17).Value = "Value"
'Cells(2, 15).Value = "Greatest % Increase"
'Cells(3, 15).Value = "Greatest % Decrease"
'Cells(4, 15).Value = "Greatest Total Volume"
'
'
'Cells(2, 16).Value = bestTick
'Cells(2, 17).Value = Format(bestVal, "Percent")
'Cells(3, 16).Value = worstTick
'Cells(3, 17).Value = Format(worstVal, "Percent")
'Cells(4, 16).Value = volTick
'Cells(4, 17).Value = volVal

'Range("Q2").Value = Format(Application.WorksheetFunction.Max(Range("K:K")), "Percent")
'Range("P2").Value = Cells(Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(Range("K:K")), Range("K:K"), 0), 9).Value
End Sub



