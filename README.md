# Stock-Ticker-VBA
HW 2
Sub MultipleYearStockData():
    Dim WorksheetName As String
    Dim i As Long
    Dim j As Long
    Dim Openrow As Long
    Dim Tickcount As Long
    Dim LastRowA As Long
    Dim LastRowI As Long
    Dim YearlyChange As Double
    Dim TotalVolume As Double
    Dim PerChange As Double
    Dim GreatIncr As Double
    Dim GreatDecr As Double
    Dim GreatVol As Double
    For Each ws In Worksheets
    ws.Activate
    Dim M As Double
    Dim N As Double
    Dim O As String
                
     
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    TotalVolume = 0
    Openrow = 2
    j = 2
    
    LastRowA = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To LastRowA
    
        TotalVolume = TotalVolume + Cells(i, "G").Value
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            ' Do calculation here'
             YearlyChange = Cells(i, "F").Value - Cells(Openrow, "C").Value
             PerChange = YearlyChange / Cells(Openrow, "C").Value
             
            ' Write values back to spreadsheet
            Range("i" & j).Value = Cells(i, 1).Value
            Range("j" & j).Value = YearlyChange
            Range("k" & j).Value = PerChange
            Range("l" & j).Value = TotalVolume
            
            If ws.Cells(j, 10).Value < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3
            Else
            ws.Cells(j, 10).Interior.ColorIndex = 4
            End If
        
            
            ' Update Variables
            
            j = j + 1
            
            If TotalVolume > GreatVol Then
                GreatVol = TotalVolume
                O = Cells(i, 1).Value
            End If
            
            
            
            TotalVolume = 0
            
            Openrow = i + 1
            
        
        End If
        
        Next i
        
    LastrowK = Cells(Rows.Count, 11).End(xlUp).Row
    MaximumValue = Application.WorksheetFunction.Max(Range("K2:K" & LastrowK))
    ws.Cells(2, 17) = MaximumValue
    M = WorksheetFunction.Match(MaximumValue, Range("K1:K" & LastrowK), 0)
    ws.Cells(2, 16) = (Range("I" & M))
    
    LastrowK = Cells(Rows.Count, 11).End(xlUp).Row
    MinimumValue = Application.WorksheetFunction.Min(Range("K2:K" & LastrowK))
    ws.Cells(3, 17) = MinimumValue
    N = WorksheetFunction.Match(MinimumValue, Range("K1:K" & LastrowK), 0)
    ws.Cells(3, 16) = (Range("I" & N))
    
    
    ws.Cells(4, 17) = GreatVol
    GreatVol = 0
    ws.Cells(4, 16) = O
    
   

    
Next ws

      
  
End Sub
