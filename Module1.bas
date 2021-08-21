Attribute VB_Name = "Module1"
Sub stockmkt()


Dim volumetotal As Double
Dim tickername As String
Dim EndYearcls As Double
Dim BegYearOpen As Double
Dim ChangeinYear As Double
Dim Percent As Double
Dim L As Double
Dim M As Double
Dim Position As Double
Dim ws As Worksheet
Dim r As Double
Dim x As Double



For Each ws In Worksheets


' I put this line just in case the files were not sorted the way they needed to be sorted to find the first open price and the cls
' price at the end of the year

ws.Range("A2").CurrentRegion.Sort key1:=ws.Range("A1"), order1:=xlAscending, key2:=ws.Range("B2"), order2:=xlAscending, Header:=xlYes

r = 2
x = 2

While ws.Cells(r, 1) <> ""
 
    volumetotal = 0
    tickername = ws.Cells(r, 1).Value
    BegYearOpen = ws.Cells(r, 3).Value
    
    ws.Cells(x, 9).Value = tickername
    
        While ws.Cells(r, 1).Value = tickername
            volumetotal = volumetotal + ws.Cells(r, 7).Value
            r = r + 1
        Wend
        
    EndYearcls = ws.Cells((r - 1), 6).Value
    
    ws.Cells(x, 12).Value = volumetotal
    ChangeinYear = EndYearcls - BegYearOpen
    ws.Cells(x, 10).Value = ChangeinYear
    
    porcent = 0
    
    If BegYearOpen = 0 And EndYearcls > 0 Then
       porcent = 1
    ElseIf BegYearOpen > 0 And EndYearcls >= 0 Then
       porcent = ChangeinYear / BegYearOpen
    End If
    
    If ChangeinYear < 0 Then
        
        colornum = 3
      Else
        
        colornum = 4
     End If
    
    ws.Cells(x, 11).Value = FormatPercent(porcent)
    ws.Cells(x, 10).Interior.ColorIndex = colornum
       
    x = x + 1
    'give the value to total
       
       
   Wend
   
   ws.Cells(1, 9).Value = "Ticker"
   ws.Cells(1, 10).Value = "Yearly Change"
   ws.Cells(1, 11).Value = "Percentage Chage"
   ws.Cells(1, 12).Value = "Total Stock Volume"
   
    L = WorksheetFunction.Max(ws.Cells.Range("K:K"))
    M = WorksheetFunction.Min(ws.Cells.Range("K:K"))
    P = WorksheetFunction.Max(ws.Cells.Range("L:L"))
    
    Position = WorksheetFunction.Match((L), ws.Cells.Range("K:K"), 0)
    Position2 = WorksheetFunction.Match((M), ws.Cells.Range("K:K"), 0)
    Position3 = WorksheetFunction.Match((P), ws.Cells.Range("L:L"), 0)
    
    ws.Cells(3, 13).Value = "Greatest % Increase"
    ws.Cells(4, 13).Value = "Greatest % Decrease"
    ws.Cells(5, 13).Value = "Greatest total vol"
    ws.Cells(3, 14).Value = ws.Cells(Position, 9).Value
    ws.Cells(4, 14).Value = ws.Cells(Position2, 9).Value
    ws.Cells(5, 14).Value = ws.Cells(Position3, 9).Value
    
    ws.Cells(3, 15).Value = FormatPercent(L)
    ws.Cells(4, 15).Value = FormatPercent(M)
    ws.Cells(5, 15).Value = P
    
    'MsgBox (L)
    'MsgBox (M)
    'MsgBox (Position)
    
   
  Next
  
End Sub
