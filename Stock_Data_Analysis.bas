Attribute VB_Name = "Module1"
Sub StockCalc()
    Dim Ticker As String
    Dim LastRow As Long
    Dim TickerRow As Long
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim Samples As Integer
    Dim finalrow As Long
    Dim firstrow As Long
    Dim CurrentValue As Double
    Dim MaxPercentChange As Double
    Dim MinPercentChange As Double
    Dim MaxTotalVolume As Double
    Dim TickerWithMaxChange As String
    Dim TickerWithMinChange As String
    Dim TickerWithMaxVolume As String
    Dim sheetName As Variant
    Dim Z As Integer

    Dim ws As Worksheet
     sheetName = Array("2018", "2019", "2020")


    For Z = 0 To 2

    ' Set the worksheet to the one you want to work with
    Set ws = ThisWorkbook.Sheets(sheetName(Z))
    
    ' Initialize the TickerRow and Ticker variables
    TickerRow = 2
    TotalVolume = 0
    finalrow = 1
    MaxPercentChange = 0.05
    MinPercentChange = 0.05
    MaxTotalVolume = 0
    TickerWithMaxChange = ""
    TickerWithMinChange = ""
    TickerWithMaxVolume = ""
    
    ' Find the last used row in column A
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
   
    For i = 2 To LastRow ' Start from the second row as you already initialized variables for the first row
        'CurrentValue = ws.Cells(i, 11).Value ' Assuming Percent Change is in column K
        Ticker = ws.Cells(i, 1).Value ' Assuming Ticker symbols are in column A
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value ' Assuming Total Volume is in column G



        ' Check for greatest total volume
        If TotalVolume > MaxTotalVolume Then
            MaxTotalVolume = TotalVolume
            TickerWithMaxVolume = Ticker
        End If

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            ws.Range("I" & TickerRow).Value = Ticker
            Samples = i - finalrow 'To find the number of samples of each ticker
            firstrow = i - Samples + 1 'To find the first row of each ticker
            OpenPrice = ws.Cells(firstrow, 3).Value
            ClosePrice = ws.Cells(i, 6).Value
            PriceChange = ClosePrice - OpenPrice
            ws.Range("J" & TickerRow).Value = PriceChange

            ' Formatting the cell
            Set PriceChangeCell = ws.Range("J" & TickerRow)
            PriceChangeCell.Value = PriceChange

            If PriceChange > 0 Then
                PriceChangeCell.Interior.Color = RGB(0, 128, 0) ' Green interior color for positive numbers
            ElseIf PriceChange < 0 Then
                PriceChangeCell.Interior.Color = RGB(255, 0, 0) ' Red interior color for negative numbers
            Else
                PriceChangeCell.Interior.Color = RGB(255, 255, 255) ' White interior color for zero
            End If

            ' Set the font color to black
            PriceChangeCell.Font.Color = RGB(0, 0, 0)

            PercentChange = PriceChange / OpenPrice
            
            If TickerRow = 2 Then
            MaxPercentChange = PercentChange
            MinPercentChange = PercentChange
            TickerWithMaxChange = Ticker
            TickerWithMinChange = Ticker
            End If
            
            ' Check for greatest percent change
            If PercentChange > MaxPercentChange Then
            MaxPercentChange = PercentChange
            TickerWithMaxChange = Ticker
            End If
        
            ' Check for greatest decrease
            If PercentChange < MinPercentChange Then
            MinPercentChange = PercentChange
            TickerWithMinChange = Ticker
            End If
            
            ws.Range("K" & TickerRow).Value = Format(PercentChange, "0.00%")
            ws.Range("L" & TickerRow).Value = TotalVolume
            finalrow = i
            TickerRow = TickerRow + 1
            TotalVolume = 0
        End If
    Next i

    ' Display the results for greatest percent change
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("P2").Value = TickerWithMaxChange
    ws.Range("O2").Value = "Greatest % Change"
    ws.Range("Q1").Value = "Value"
    ws.Range("Q2").Value = MaxPercentChange * 100#

    ws.Range("P3").Value = TickerWithMinChange
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("Q3").Value = MinPercentChange * 100#
    
    ws.Range("P4").Value = TickerWithMaxVolume
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("Q4").Value = MaxTotalVolume
    
    Next Z
End Sub
