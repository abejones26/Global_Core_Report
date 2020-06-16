Sub CoreReport()


Dim i As Long
Dim lastrow As Integer
Dim open_qty As Double
Dim open_qty_total As Double
Dim month_receipts_total As Long
Dim percentage As Double
Dim balance As Long
Dim po As String
Dim sTemp As String
Dim ws As Worksheet
Dim space As String

Set ws = Sheets("CoreReportAll")
open_qty = 0
open_qty_total = 0
month_receipts_total = 0
balance = 0
percentage = 0

lastrow = Cells(Rows.Count, 2).End(xlUp).Row

' Adding Duplicate '

    Range("A3").Select
    ActiveCell.FormulaR1C1 = "=IF(R[-1]C[1]=RC[1],""Duplicate"","""")"
    Selection.AutoFill Destination:=Range("A3:A10000")
    Range("A3:A10000").Select
    Range("A1").Select
    ActiveCell.FormulaR1C1 = ""
    Range("A1").Select
    
' First month dates
    
    Range("T1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(WEEKDAY(EOMONTH(TODAY(),R3C13-1))=1,WEEKDAY(EOMONTH(TODAY(),R3C13-1))=2),((EOMONTH(TODAY(),R3C13-1))-WEEKDAY(EOMONTH(TODAY(),R3C13-1)))-14,((EOMONTH(TODAY(),R3C13-1))+6-MOD((EOMONTH(TODAY(),R3C13-1))-1,7))-14)"
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]-WEEKDAY(RC[1],1)"
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]-WEEKDAY(RC[1],1)"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]-WEEKDAY(RC[1],1)"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]-WEEKDAY(RC[1],1)"
    
' Second month dates

    Range("AD1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(WEEKDAY(EOMONTH(TODAY(),R3C13))=1,WEEKDAY(EOMONTH(TODAY(),R3C13))=2),((EOMONTH(TODAY(),R3C13))-WEEKDAY(EOMONTH(TODAY(),R3C13)))-14,((EOMONTH(TODAY(),R3C13))+6-MOD((EOMONTH(TODAY(),R3C13))-1,7))-14)"
    Range("AC1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]-WEEKDAY(RC[1],1)"
    Range("AB1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]-WEEKDAY(RC[1],1)"
    Range("AA1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]-WEEKDAY(RC[1],1)"
    Range("Z1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((RC[1]-WEEKDAY(RC[1],1))=RC[-6],"""",(RC[1]-WEEKDAY(RC[1],1)))"
    Range("Y1").Select
    
' Third month dates

    Range("AN1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(WEEKDAY(EOMONTH(TODAY(),R3C13+1))=1,WEEKDAY(EOMONTH(TODAY(),R3C13+1))=2),((EOMONTH(TODAY(),R3C13+1))-WEEKDAY(EOMONTH(TODAY(),R3C13+1)))-14,((EOMONTH(TODAY(),R3C13+1))+6-MOD((EOMONTH(TODAY(),R3C13+1))-1,7))-14)"
    Range("AM1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]-WEEKDAY(RC[1],1)"
    Range("AL1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]-WEEKDAY(RC[1],1)"
    Range("AK1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]-WEEKDAY(RC[1],1)"
    Range("AJ1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((RC[1]-WEEKDAY(RC[1],1))=RC[-6],"""",(RC[1]-WEEKDAY(RC[1],1)))"
    Range("AI1").Select
    
' Fourth month dates

    Range("AX1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(WEEKDAY(EOMONTH(TODAY(),R3C13+2))=1,WEEKDAY(EOMONTH(TODAY(),R3C13+2))=2),((EOMONTH(TODAY(),R3C13+2))-WEEKDAY(EOMONTH(TODAY(),R3C13+2)))-14,((EOMONTH(TODAY(),R3C13+2))+6-MOD((EOMONTH(TODAY(),R3C13+2))-1,7))-14)"
    Range("AW1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]-WEEKDAY(RC[1],1)"
    Range("AV1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]-WEEKDAY(RC[1],1)"
    Range("AU1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]-WEEKDAY(RC[1],1)"
    Range("AT1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((RC[1]-WEEKDAY(RC[1],1))=RC[-6],"""",(RC[1]-WEEKDAY(RC[1],1)))"
    Range("AS1").Select
    
' Fifth month dates

    Range("BH1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(WEEKDAY(EOMONTH(TODAY(),R3C13+3))=1,WEEKDAY(EOMONTH(TODAY(),R3C13+3))=2),((EOMONTH(TODAY(),R3C13+3))-WEEKDAY(EOMONTH(TODAY(),R3C13+3)))-14,((EOMONTH(TODAY(),R3C13+3))+6-MOD((EOMONTH(TODAY(),R3C13+3))-1,7))-14)"
    Range("BG1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]-WEEKDAY(RC[1],1)"
    Range("BF1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]-WEEKDAY(RC[1],1)"
    Range("BE1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]-WEEKDAY(RC[1],1)"
    Range("BD1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((RC[1]-WEEKDAY(RC[1],1))=RC[-6],"""",(RC[1]-WEEKDAY(RC[1],1)))"
    Range("BC1").Select
    
' Six month dates

    Range("BR1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(WEEKDAY(EOMONTH(TODAY(),R3C13+4))=1,WEEKDAY(EOMONTH(TODAY(),R3C13+4))=2),((EOMONTH(TODAY(),R3C13+4))-WEEKDAY(EOMONTH(TODAY(),R3C13+4)))-14,((EOMONTH(TODAY(),R3C13+4))+6-MOD((EOMONTH(TODAY(),R3C13+4))-1,7))-14)"
    Range("BQ1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]-WEEKDAY(RC[1],1)"
    Range("BP1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]-WEEKDAY(RC[1],1)"
    Range("BO1").Select
    ActiveCell.FormulaR1C1 = "=RC[1]-WEEKDAY(RC[1],1)"
    Range("BN1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF((RC[1]-WEEKDAY(RC[1],1))=RC[-6],"""",(RC[1]-WEEKDAY(RC[1],1)))"
    Range("BM1").Select

    ' Removing code for weekly dates

    Range("P1:BR1").Select
    Range("P1:BR1").Copy
    Range("P1:BR1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
' Removing code for Duplicate column

    Columns("A:A").Select
    Columns("A:A").Copy
    Columns("A:A").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

' rcptmonth Macro

    Range("CD2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(WEEKDAY(EOMONTH(RC[5],0))=1,WEEKDAY(EOMONTH(RC[5],0))=2),((EOMONTH(RC[5],0))-WEEKDAY(EOMONTH(RC[5],0)))-14,((EOMONTH(RC[5],0))+6-MOD((EOMONTH(RC[5],0))-1,7))-14)"
    Range("CD2").Select
    Selection.AutoFill Destination:=Range("CD2:CD10000")
    Range("CD2:CD10000").Select
    Columns("CD:CD").Select
    Columns("CD:CD").Copy
    Columns("CD:CD").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("CD:CD").Select
    Selection.NumberFormat = "[$-en-US]d-mmm;@"

'Adding current ship date

    Range("CH2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]>0,RC[-1],IF(RC[-2]=RC[-3],"""",RC[-2]))"
    Range("CH2").Select
    Selection.AutoFill Destination:=Range("CH2:CH10000")
    Range("CH2:CH10000").Select
    Columns("CH:CH").Copy
    Columns("CH:CH").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

' Conditional Format for percentage columns

    Range("W:W,AG:AG,AQ:AQ").Select
    Range("AQ1").Activate
    Range("W:W,AG:AG,AQ:AQ,BA:BA,BK:BK,BU:BU").Select
    Range("BU1").Activate
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual _
        , Formula1:="=3.5"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLessEqual, _
        Formula1:="=1"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("A1").Select

' Using For loop to enter open units into weekly buckets

For i = 2 To 11000
    ' Solving for Week 1
    If Cells(i, 87).Value <= Cells(1, 16).Value Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 16).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 2
    If Cells(i, 87).Value > Cells(1, 16).Value And Cells(i, 87).Value <= Cells(1, 17) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 17).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 3
    If Cells(i, 87).Value > Cells(1, 17).Value And Cells(i, 87).Value <= Cells(1, 18) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 18).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 4
    If Cells(i, 87).Value > Cells(1, 18).Value And Cells(i, 87).Value <= Cells(1, 19) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 19).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 5
    If Cells(i, 87).Value > Cells(1, 19).Value And Cells(i, 87).Value <= Cells(1, 20) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 20).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 6
    If Cells(1, 26).Value = "" Then
        open_qty = 0
    ElseIf Cells(i, 87).Value > Cells(1, 20).Value And Cells(i, 87).Value <= Cells(1, 26) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 26).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 7
    If Cells(1, 26).Value = "" Then
        If Cells(i, 87).Value > Cells(1, 20).Value And Cells(i, 87).Value <= Cells(1, 27) Then
            open_qty = open_qty + Cells(i, 81).Value
            Cells(i, 27).Value = open_qty
            open_qty = 0
        End If
    ElseIf Cells(i, 87).Value > Cells(1, 26).Value And Cells(i, 87).Value <= Cells(1, 27) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 27).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 8
    If Cells(i, 87).Value > Cells(1, 27).Value And Cells(i, 87).Value <= Cells(1, 28) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 28).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 9
    If Cells(i, 87).Value > Cells(1, 28).Value And Cells(i, 87).Value <= Cells(1, 29) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 29).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 10
    If Cells(i, 87).Value > Cells(1, 29).Value And Cells(i, 87).Value <= Cells(1, 30) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 30).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 11
    If Cells(1, 36).Value = "" Then
        open_qty = 0
    ElseIf Cells(i, 87).Value > Cells(1, 30).Value And Cells(i, 87).Value <= Cells(1, 36) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 36).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 12
    If Cells(1, 36).Value = "" Then
        If Cells(i, 87).Value > Cells(1, 30).Value And Cells(i, 87).Value <= Cells(1, 37) Then
            open_qty = open_qty + Cells(i, 81).Value
            Cells(i, 37).Value = open_qty
            open_qty = 0
        End If
    ElseIf Cells(i, 87).Value > Cells(1, 36).Value And Cells(i, 87).Value <= Cells(1, 37) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 37).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 13
    If Cells(i, 87).Value > Cells(1, 37).Value And Cells(i, 87).Value <= Cells(1, 38) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 38).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 14
    If Cells(i, 87).Value > Cells(1, 38).Value And Cells(i, 87).Value <= Cells(1, 39) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 39).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 15
    If Cells(i, 87).Value > Cells(1, 39).Value And Cells(i, 87).Value <= Cells(1, 40) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 40).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 16
    If Cells(1, 46).Value = "" Then
        open_qty = 0
    ElseIf Cells(i, 87).Value > Cells(1, 40).Value And Cells(i, 87).Value <= Cells(1, 46) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 46).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 17
    If Cells(1, 46).Value = "" Then
        If Cells(i, 87).Value > Cells(1, 40).Value And Cells(i, 87).Value <= Cells(1, 47) Then
            open_qty = open_qty + Cells(i, 81).Value
            Cells(i, 47).Value = open_qty
            open_qty = 0
        End If
    ElseIf Cells(i, 87).Value > Cells(1, 46).Value And Cells(i, 87).Value <= Cells(1, 47) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 47).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 18
    If Cells(i, 87).Value > Cells(1, 47).Value And Cells(i, 87).Value <= Cells(1, 48) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 48).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 19
    If Cells(i, 87).Value > Cells(1, 48).Value And Cells(i, 87).Value <= Cells(1, 49) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 49).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 20
    If Cells(i, 87).Value > Cells(1, 49).Value And Cells(i, 87).Value <= Cells(1, 50) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 50).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 21
    If Cells(1, 56).Value = "" Then
        open_qty = 0
    ElseIf Cells(i, 87).Value > Cells(1, 50).Value And Cells(i, 87).Value <= Cells(1, 56) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 56).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 22
    If Cells(1, 56).Value = "" Then
        If Cells(i, 87).Value > Cells(1, 50).Value And Cells(i, 87).Value <= Cells(1, 57) Then
            open_qty = open_qty + Cells(i, 81).Value
            Cells(i, 57).Value = open_qty
            open_qty = 0
        End If
    ElseIf Cells(i, 87).Value > Cells(1, 56).Value And Cells(i, 87).Value <= Cells(1, 57) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 57).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 23
    If Cells(i, 87).Value > Cells(1, 57).Value And Cells(i, 87).Value <= Cells(1, 58) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 58).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 24
    If Cells(i, 87).Value > Cells(1, 58).Value And Cells(i, 87).Value <= Cells(1, 59) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 59).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 25
    If Cells(i, 87).Value > Cells(1, 59).Value And Cells(i, 87).Value <= Cells(1, 60) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 60).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 26
    If Cells(1, 66).Value = "" Then
        open_qty = 0
    ElseIf Cells(i, 87).Value > Cells(1, 60).Value And Cells(i, 87).Value <= Cells(1, 66) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 66).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 27
    If Cells(1, 66).Value = "" Then
        If Cells(i, 87).Value > Cells(1, 60).Value And Cells(i, 87).Value <= Cells(1, 67) Then
            open_qty = open_qty + Cells(i, 81).Value
            Cells(i, 67).Value = open_qty
            open_qty = 0
        End If
    ElseIf Cells(i, 87).Value > Cells(1, 66).Value And Cells(i, 87).Value <= Cells(1, 67) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 67).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 28
    If Cells(i, 87).Value > Cells(1, 67).Value And Cells(i, 87).Value <= Cells(1, 68) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 68).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 29
    If Cells(i, 87).Value > Cells(1, 68).Value And Cells(i, 87).Value <= Cells(1, 69) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 69).Value = open_qty
        open_qty = 0
    End If
Next i

For i = 2 To 11000
    ' Solving for Week 30
    If Cells(i, 87).Value > Cells(1, 69).Value And Cells(i, 87).Value <= Cells(1, 70) Then
        open_qty = open_qty + Cells(i, 81).Value
        Cells(i, 70).Value = open_qty
        open_qty = 0
    End If
Next i

' Adding POs into monthly buckets

For i = 2 To 11000
    ' Month 1 POs
    If Cells(i, 87).Value <= Cells(1, 20).Value Then
        po = Cells(i, 77).Value
        Cells(i, 15).Value = po
        po = ""
    End If
Next i

For i = 2 To 11000
    ' Month 2 POs
    po = ""
    If Cells(i, 87).Value > Cells(1, 20).Value And Cells(i, 87).Value <= Cells(1, 30).Value Then
        po = Cells(i, 77).Value
        Cells(i, 25).Value = po
        po = ""
    End If
Next i

For i = 2 To 11000
    ' Month 3 POs
    po = ""
    If Cells(i, 87).Value > Cells(1, 30).Value And Cells(i, 87).Value <= Cells(1, 40).Value Then
        po = Cells(i, 77).Value
        Cells(i, 35).Value = po
        po = ""
    End If
Next i

For i = 2 To 11000
    ' Month 4 POs
    po = ""
    If Cells(i, 87).Value > Cells(1, 40).Value And Cells(i, 87).Value <= Cells(1, 50).Value Then
        po = Cells(i, 77).Value
        Cells(i, 45).Value = po
        po = ""
    End If
Next i

For i = 2 To 11000
    ' Month 5 POs
    po = ""
    If Cells(i, 87).Value > Cells(1, 50).Value And Cells(i, 87).Value <= Cells(1, 60).Value Then
        po = Cells(i, 77).Value
        Cells(i, 55).Value = po
        po = ""
    End If
Next i

For i = 2 To 11000
    ' Month 6 POs
    po = ""
    If Cells(i, 87).Value > Cells(1, 60).Value And Cells(i, 87).Value <= Cells(1, 70).Value Then
        po = Cells(i, 77).Value
        Cells(i, 65).Value = po
        po = ""
    End If
Next i

' Adding totals row

For i = 3 To 11000
        If i = 3 Then
            sTemp = ws.Cells(i, 2).Value
        Else
            If ws.Cells(i, 2).Value <> sTemp Then
                sTemp = ws.Cells(i, 2).Value
                ws.Rows(i).EntireRow.Insert
                ws.Rows(i).EntireRow.Interior.Pattern = xlSolid
                ws.Rows(i).EntireRow.Interior.PatternColorIndex = xlAutomatic
                ws.Rows(i).EntireRow.Interior.ThemeColor = xlThemeColorAccent1
                ws.Rows(i).EntireRow.Interior.TintAndShade = 0.799981688894314
                ws.Rows(i).EntireRow.Interior.PatternTintAndShade = 0
            End If
        End If
    Next i

' Filling in blanks range B:N with its style values

    Columns("B:N").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.FormulaR1C1 = "=R[-1]C"
    Columns("B:N").Select
    Columns("B:N").Copy
    Columns("B:N").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Columns("U:X").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.FormulaR1C1 = "=R[-1]C"
    Columns("U:X").Select
    Columns("U:X").Copy
    Columns("U:X").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Columns("AE:AH").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.FormulaR1C1 = "=R[-1]C"
    Columns("AE:AH").Select
    Columns("AE:AH").Copy
    Columns("AE:AH").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    Columns("AO:AR").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.FormulaR1C1 = "=R[-1]C"
    Columns("AO:AR").Select
    Columns("AO:AR").Copy
    Columns("AO:AR").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    Columns("AY:BB").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.FormulaR1C1 = "=R[-1]C"
    Columns("AY:BB").Select
    Columns("AY:BB").Copy
    Columns("AY:BB").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Columns("BI:BL").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.FormulaR1C1 = "=R[-1]C"
    Columns("BI:BL").Select
    Columns("BI:BL").Copy
    Columns("BI:BL").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Columns("BS:BW").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.FormulaR1C1 = "=R[-1]C"
    Columns("BS:BW").Select
    Columns("BS:BW").Copy
    Columns("BS:BW").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Columns("CM").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Selection.FormulaR1C1 = "=R[-1]C"
    Columns("CM").Select
    Columns("CM").Copy
    Columns("CM").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    