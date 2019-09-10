Attribute VB_Name = "Module11"
Sub stock():

Dim tickerName As String

'Dim i As Integer

Dim LastRow As Long
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

Dim TotalStockVolume As Double

Dim Summary_Table_Row As Integer

Dim OpenPrice As Double

Dim ClosePrice As Double

Dim YearlyChangeInPrice As Double

Dim GreatestPercentIncrease As Double

Dim GreatestPercentIncreaseTicker As String


Dim GreatestPercentDecrease As Double
Dim GreatestPercentDecreaseTicker As String

Dim changeInPricePct As Double

Dim GreatestTotalSum As Double
Dim GreatestTotalSumTicker As String

Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning

For Each ws In ThisWorkbook.Worksheets
ws.Activate
TotalStockVolume = 0#
Summary_Table_Row = 2
GreatestPercentDecrease = 0#
GreatestTotalSum = 0#
GreatestPercentIncrease = 0#
YearlyChangeInPrice = 0#
OpenPrice = -1#
ClosePrice = 0#
changeInPricePct = 0#
    
For i = 2 To LastRow

        If OpenPrice = -1# Then
            OpenPrice = Cells(i, 3).Value
        End If
    
        If (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
            
            'Set the tickerName
            tickerName = Cells(i, 1).Value
            
            'Set the Total Stock Volume
            TotalStockVolume = TotalStockVolume# + Cells(i, 7).Value#
        
            'Initialize the tickerName with the unique TickerName
            'until the new Ticker name is found and add it to Summary table Row
            Cells(Summary_Table_Row, 9).Value = tickerName
            
            'Print the Total Stock Volume
            Cells(Summary_Table_Row, 12).Value = TotalStockVolume#
        
            
            'Reset the Total Stock Volume to Zero for the next Ticker
            
            ClosePrice = Cells(i, 6).Value
            
            YearlyChangeInPrice = ClosePrice - OpenPrice
            
            'Print the YearlyChangeinPrice
            Cells(Summary_Table_Row, 10).Value = YearlyChangeInPrice
            
            'Print the Percentage Change in Price
            If OpenPrice = 0 Then
                Cells(Summary_Table_Row, 11).Value = "NAN"
                Cells(Summary_Table_Row, 11).Interior.ColorIndex = 5
            
            Else
                changeInPricePct = Round((YearlyChangeInPrice / OpenPrice) * 100, 2)
                Cells(Summary_Table_Row, 11).Value = changeInPricePct
                If changeInPricePct >= 0 Then
                    Cells(Summary_Table_Row, 11).Interior.ColorIndex = 4
                ElseIf changeInPricePct < 0 Then
                    Cells(Summary_Table_Row, 11).Interior.ColorIndex = 3
                End If
                If changeInPricePct > GreatestPercentIncrease Then
                    GreatestPercentIncrease = changeInPricePct
                    GreatestPercentIncreaseTicker = tickerName
                End If
                
                If changeInPricePct < GreatestPercentDecrease Then
                    GreatestPercentDecrease = changeInPricePct
                    GreatestPercentDecreaseTicker = tickerName
                End If
            
            End If
            
            If TotalStockVolume > GreatestTotalSum Then
                GreatestTotalSum = TotalStockVolume
                GreatestTotalSumTicker = tickerName
            End If
           
            TotalStockVolume = 0#
            
            'Print to the next row of the Summary table
            Summary_Table_Row = Summary_Table_Row + 1
         
            OpenPrice = -1#
            ClosePrice = 0#
        Else
        
            'Add to the previous Total Stock Volume for the same Ticker
            'cell_value = Cells(i, 7).Value#
            TotalStockVolume = TotalStockVolume# + Cells(i, 7).Value#
            
        End If
    
Next i


'Print Greatest of Percent Increase, Percent Decrease and Total Sum to Column 15
            Cells(2, 13).Value = "GreatestPercentIncrease"
            Cells(3, 13).Value = "GreatestPercentDecrease"
            Cells(4, 13).Value = "GreatestTotalSum"
            Cells(2, 14).Value = GreatestPercentIncreaseTicker
            Cells(3, 14).Value = GreatestPercentDecreaseTicker
            Cells(4, 14).Value = GreatestTotalSumTicker
            Cells(2, 15).Value = GreatestPercentIncrease
            Cells(3, 15).Value = GreatestPercentDecrease
            Cells(4, 15).Value = GreatestTotalSum
            
            'DO NOT USE - Reset the Yearly Change in Price for the next Ticker
            'YearlyChangeInPrice = 0#
Next

starting_ws.Activate 'activate the worksheet that was originally active
            
           
End Sub

