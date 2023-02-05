Attribute VB_Name = "Module1"
Sub stock()

    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Activate
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
    
        lastrowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Dim stocknameA As String
        Dim stotalA As Double
        stotalA = 0
        Dim srowA As Long
        srowA = 2
        Dim lastvalueA As Double
        Dim firstvalueA As Double
        Dim difA As Double
        Dim grabfirstA As Double
        Dim grablastA As Double
        Dim percentA As Double
        Dim maxper, maxperIndex As Double
        Dim minper, minperIndex As Double
        Dim maxvol, maxvolIndex As Integer
    
        Dim rA As Long
        For rA = 2 To lastrowA
            If ws.Cells(rA + 1, 1).Value <> ws.Cells(rA, 1).Value Then
                ' MsgBox (ws.Cells(rA, 1).Value + " -> " + ws.Cells(rA + 1, 1).Value)
                ' MsgBox (Worksheets("A").Cells(rA + 1, 3).Value)
                stocknameA = ws.Cells(rA, 1).Value
                lastvalueA = ws.Cells(rA, 6).Value
                firstvalueA = ws.Cells(rA + 1, 3).Value
                stotalA = stotalA + ws.Cells(rA, 7).Value
                ws.Cells(srowA, 9).Value = stocknameA
                ws.Cells(srowA, 12).Value = stotalA
                ws.Cells(srowA + 1, 19).Value = firstvalueA
                ws.Cells(srowA, 20).Value = lastvalueA
                ws.Cells(2, 19).Value = ws.Cells(2, 3).Value
                grabfirstA = ws.Cells(srowA, 19).Value
                grablastA = ws.Cells(srowA, 20).Value
                difA = grablastA - grabfirstA
                ws.Cells(srowA, 10).Value = difA
                percentA = ((grablastA / grabfirstA) - 1)
                ws.Cells(srowA, 11).Value = percentA
                ws.Cells(srowA, 11).NumberFormat = "0.00%"
    
                If ws.Cells(srowA, 10).Value >= 0 Then
                    ws.Cells(srowA, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(srowA, 10).Interior.ColorIndex = 3
                End If
    
                srowA = srowA + 1
                stotalA = 0
            Else
                stotalA = stotalA + ws.Cells(rA, 7).Value
            End If
        Next rA
    
        maxper = WorksheetFunction.Max(ws.Range("K:K"))
        maxperIndex = WorksheetFunction.Match(maxper, ws.Range("K:K"), 0)
        ws.Range("P2").Value = ws.Range("I" & maxperIndex).Value
        ws.Range("Q2").Value = maxper
        ws.Range("Q2").NumberFormat = "0.00%"
        minper = WorksheetFunction.Min(ws.Range("K:K"))
        minperIndex = WorksheetFunction.Match(minper, ws.Range("K:K"), 0)
        ws.Range("P3").Value = ws.Range("I" & minperIndex).Value
        ws.Range("Q3").Value = minper
        ws.Range("Q3").NumberFormat = "0.00%"
        maxvol = WorksheetFunction.Max(ws.Range("L:L"))
        maxvolIndex = WorksheetFunction.Match(maxvol, ws.Range("L:L"), 0)
        ws.Range("P4").Value = ws.Range("I" & maxvolIndex).Value
        ws.Range("Q4").Value = maxvol
    
            ws.Range("J:L").Columns.AutoFit
            ws.Range("O:O").Columns.AutoFit
            ws.Range("Q:Q").Columns.AutoFit
    
        Dim first As Double
    
        ws.Range("S:S").Value = ""
        ws.Range("T:T").Value = ""

    Next

End Sub
