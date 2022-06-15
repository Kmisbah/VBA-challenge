
Sub TickerLoop()  
    Dim WS As Worksheet
    Application.ScreenUpdating = False
    
    For Each WS In Worksheets
        WS.Select
        Call TickerTape
    Next
    
    Application.ScreenUpdating = True
End Sub


Sub TickerTape():

Dim i As Double 
Dim j As Double 

Dim tickerALL As Double 
Dim opening As Double 
Dim closing As Double 
Dim volume As Double 

Dim SUMMline As Double 
Dim tickerSUMM As Double 
Dim YearChgCol As Double 
Dim PerChgCol As Double 
Dim TotalStockCol As Double 

Dim TickerAn As Double 
Dim TickerSym As Double 
Dim Value As Double 

Dim yearchange As Double 
Dim perchange As Double 
Dim tot_ticker As Double 
Dim totalstock As Double 
Dim lastrow As Double 
Dim firstline As Double 
Dim lastline As Double

Dim MaxInc As Double 
Dim MaxDec As Double
Dim MaxTotVol As Double 
Dim FindRow As Double 


tickerALL = 1
opening = 3
closing = 6
volume = 7

SUMMline = 1
tickerSUMM = 9
YearChgCol = 10
PerChgCol = 11
TotalStockCol = 12

TickerAn = 15
TickerSym = TickerAn + 1
Value = TickerAn + 2


tot_tickers = 1
totalstock = 0
firstline = 2

        
lastrow = Cells(Rows.Count, tickerALL).End(xlUp).Row



Cells(SUMMline + 1, tickerSUMM).Value = Cells(2, tickerALL)

For i = 2 To lastrow
    totalstock = totalstock + Cells(i, volume)
    If Cells(i + 1, tickerALL).Value <> Cells(i, tickerALL).Value Then
        lastline = i
        
        yearchange = Cells(lastline, closing).Value - Cells(firstline, opening).Value
        Cells(tot_tickers + 1, YearChgCol) = yearchange
        
      
        If Cells(firstline, opening).Value = 0 Then
            Cells(tot_tickers + 1, PerChgCol) = "NaN"
        Else
            perchange = yearchange / Cells(firstline, opening).Value
            Cells(tot_tickers + 1, PerChgCol) = perchange
        End If
        
          
        Cells(tot_tickers + 1, TotalStockCol) = totalstock
        
        
        totalstock = 0
        firstline = i + 1
        Cells(tot_tickers + 2, tickerSUMM).Value = Cells(i + 1, tickerALL) 
        tot_tickers = tot_tickers + 1
    End If
Next i
      
        

Cells(SUMMline, tickerSUMM).Value = "Ticker"

Cells(SUMMline, YearChgCol).Value = "Yearly Change"

    Cells(SUMMline + 1, YearChgCol).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual _
        , Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 3407718
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Columns(YearChgCol).EntireColumn.AutoFit

Cells(SUMMline, PerChgCol).Value = "Percent Change"
    Columns(PerChgCol).NumberFormat = "0.00%"
    Columns(PerChgCol).EntireColumn.AutoFit
Cells(SUMMline, TotalStockCol).Value = "Total Stock Volume"
    Columns(TotalStockCol).EntireColumn.AutoFit
 
      

End Sub
