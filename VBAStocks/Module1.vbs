Sub main1()

Application.ScreenUpdating = True
WS_count = ActiveWorkbook.Worksheets.Count

' loop through all sheets
    For i = 1 To WS_count
                
    'activate current ws
    Worksheets(i).Activate
    
    'clear everything in cols I:M
    Columns("I:Q").Select
    Selection.Delete Shift:=xlToLeft
    Selection.FormatConditions.Delete
    Cells.FormatConditions.Delete
    
    'Ticker
    Call ticker
    
    'yearly change
    Call yearlychange2
    
    'call total volume
    Call totalvolume
    
    'find highest changers
    Call highestchange
    
    'zoom set
    ActiveWindow.Zoom = 100
    
    'autofit columns
    Columns("I:Q").Select
    Columns("I:Q").EntireColumn.AutoFit
    
    'activate a1 again just for cleanliness
    Range("a1").Activate
        
    Next i

Range("a1").Activate


End Sub


Sub ticker()

    Dim ticker() As String
    
    Range("a2").Activate
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("I2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.RemoveDuplicates Columns:=1, Header:= _
        xlNo
    
    Range("i1").Value = "Ticker"

End Sub


Sub totalvolume()

   Dim lastrow As Long
   lastrow = Range("K" & Rows.Count).End(xlUp).Row
   Range("L2:L" & lastrow).Formula = "=SUMIFS($G:$G,$A:$A,I2)"
      
   Range("L1").Value = "Total Stock Volume"

End Sub

Sub highestchange()
    
    'get the values
    Range("k2").Activate
    Range(Selection, Selection.End(xlDown)).Select
    
    h_increase = WorksheetFunction.Max(Selection)
    ticker_increase = WorksheetFunction.Match(h_increase, Selection, 0)
    
    h_decrease = WorksheetFunction.Min(Selection)
    ticker_decrease = WorksheetFunction.Match(h_decrease, Selection, 0)
    
    Range("L2").Activate
    Range(Selection, Selection.End(xlDown)).Select
    
    h_volume = WorksheetFunction.Max(Selection)
    ticker_volume = WorksheetFunction.Match(h_volume, Selection, 0)
   
    
    
    'paste in the right cells
    Range("o2").Value = "Greatest % increase"
    Range("o3").Value = "Greatest % decrease"
    Range("o4").Value = "Greatest Total Volume"
    
    
    Range("p1").Value = "Ticker"
    Range("i2").Activate
    Range(Selection, Selection.End(xlDown)).Select
    Range("p2").Value = WorksheetFunction.Index(Selection, ticker_increase, 1)
    Range("p3").Value = WorksheetFunction.Index(Selection, ticker_decrease, 1)
    Range("p4").Value = WorksheetFunction.Index(Selection, ticker_volume, 1)
    
    Range("q1").Value = "Value"
    Range("q2").Value = h_increase
    Range("q3").Value = h_decrease
    Range("q4").Value = h_volume
    
    Range("Q2:Q3").Select
    Selection.NumberFormat = "0.00%"
    
    Range("Q4").Select
    Selection.NumberFormat = "General"
    
    
    
    
    
End Sub





Sub yearlychange2()
    Columns("J:q").Select
    Selection.Delete Shift:=xlToLeft
    
    Dim tickerlist() As Variant
    Dim tickerlist_full() As Variant
    Dim dates1() As Variant
    Dim openprice() As Variant
    Dim closeprice() As Variant
    
    'ticker import
    Range("i2").Activate
    tickerlist = Range(Selection, Selection.End(xlDown))
    
    'ticker import - full list
    Range("a2").Activate
    tickerlist_full = Range(Selection, Selection.End(xlDown))
        
    'dates import
    Range("b2").Activate
    dates1 = Range(Selection, Selection.End(xlDown))
    
    'openprice import
    Range("c2").Activate
    openprice = Range(Selection, Selection.End(xlDown))

    'close price import
    Range("f2").Activate
    closeprice = Range(Selection, Selection.End(xlDown))
    
    Dim lastrow As Long
    lastrow = Range("I" & Rows.Count).End(xlUp).Row
    Range("L2:L" & lastrow).Formula = "=COUNTIFS($A:$A,I2)"
    Range("M2:M" & lastrow).Formula = "=sum(l$2:l2)"
    
    
    'loop through ticker array
     For i = 1 To UBound(tickerlist)
     
     If i = 1 Then
        Cells(i + 1, 10).Value = closeprice(Cells(i + 1, 13).Value, 1) - openprice(1, 1)
        Cells(i + 1, 14).Value = openprice(1, 1)
        
        GoTo A
     End If
     
     
     Cells(i + 1, 10).Value = closeprice(Cells(i + 1, 13).Value, 1) - openprice(Cells(i, 13).Value + 1, 1)
     Cells(i + 1, 14).Value = openprice(Cells(i, 13).Value + 1, 1)
A:
     Next i
     
     Range("k2:k" & lastrow).Formula = "=iferror((J2/N2),0)"
     
     
     
    'conditional format, paste formulas as values
    Range("J2:J" & lastrow).Activate
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.NumberFormat = "General"
    

    Range("J1").Value = "Yearly change"
    
    'percent change
    Range("K2:K" & lastrow).Activate
    
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.NumberFormat = "0.00%"
    
    
    Columns("L:Q").Select
    Selection.Delete Shift:=xlToLeft
    
    Range("K1").Value = "Percent change"
     




End Sub