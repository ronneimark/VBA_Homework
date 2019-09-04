Attribute VB_Name = "Module1"
Sub StockAnalysis()
Attribute StockAnalysis.VB_ProcData.VB_Invoke_Func = "R\n14"
    
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

ws.Activate
    
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Volume"
    
    Dim Ticker As String
    Dim TotalVolume As Double
    Dim Starting As Double
    Dim Closing As Double
    Dim Percent As Double
    
    i = 2
    t = 2
    
    TotalVolume = 0
    
    Do While Cells(i, 1) <> ""
        
        Starting = Cells(i, 3).Value
        
        Do
    
            TotalVolume = TotalVolume + Cells(i, 7).Value
            i = i + 1
            
        Loop While Cells(i, 1).Value = Cells(i - 1, 1).Value
    
        Closing = Cells(i - 1, 6).Value
        
        Cells(t, 9).Value = Cells(i - 1, 1)
        Cells(t, 10).Value = Closing - Starting
        If Cells(t, 10).Value <= 0 Then Cells(t, 10).Interior.Color = vbRed
        If Cells(t, 10).Value > 0 Then Cells(t, 10).Interior.Color = vbGreen
        
        If Starting = 0 Then
            Cells(t, 11).Value = 0
        Else
            Cells(t, 11).Value = FormatPercent((Closing - Starting) / Starting, [2])
        End If
        
        Cells(t, 12).Value = TotalVolume

        
        TotalVolume = 0
        t = t + 1
    
    Loop
    
    Cells(2, 14) = "Greatest % Increase"
    Cells(3, 14) = "Greatest % Decrease"
    Cells(4, 14) = "Greatest Total Volume"
    Cells(1, 15) = "Ticker"
    Cells(1, 16) = "Value"
    
    Dim HighestPercent As Double
    Dim LowestPercent As Double
    Dim GreatestVolume As Double
    Dim HighestTicker As String
    Dim LowestTicker As String
    Dim HighVolTicker As String
    

    j = 2

    Do While Cells(j, 9) <> ""
    
    
        If Cells(j, 11).Value > HighestPercent Then
            HighestPercent = Cells(j, 11).Value
            HighestTicker = Cells(j, 9).Value
            
        End If
            
        If Cells(j, 11).Value < LowestPercent Then
            LowestPercent = Cells(j, 11).Value
            LowestTicker = Cells(j, 9).Value
        
        End If
            
        If Cells(j, 12).Value > GreatestVolume Then
            GreatestVolume = Cells(j, 12).Value
            HighVolTicker = Cells(j, 9).Value
            
        End If
        
        j = j + 1
            
    Loop
     
    Cells(2, 15).Value = HighestTicker
    Cells(2, 16).Value = FormatPercent(HighestPercent, [2])
    Cells(3, 15).Value = LowestTicker
    Cells(3, 16).Value = FormatPercent(LowestPercent, [2])
    Cells(4, 15).Value = HighVolTicker
    Cells(4, 16).Value = GreatestVolume
  
Next ws

End Sub
