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

' Week 1 total

For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 16).Value
        Cells(i, 16).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 16).Value
    End If
Next i

' Week 2 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 17).Value
        Cells(i, 17).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 17).Value
    End If
Next i

' Week 3 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 18).Value
        Cells(i, 18).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 18).Value
    End If
Next i

' Week 4 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 19).Value
        Cells(i, 19).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 19).Value
    End If
Next i

' Week 5 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 20).Value
        Cells(i, 20).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 20).Value
    End If
Next i

' Week 6 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 26).Value
        Cells(i, 26).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 26).Value
    End If
Next i

' Week 7 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 27).Value
        Cells(i, 27).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 27).Value
    End If
Next i

' Week 8 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 28).Value
        Cells(i, 28).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 28).Value
    End If
Next i

' Week 9 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 29).Value
        Cells(i, 29).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 29).Value
    End If
Next i

' Week 10 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 30).Value
        Cells(i, 30).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 30).Value
    End If
Next i

' Week 11 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 36).Value
        Cells(i, 36).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 36).Value
    End If
Next i

' Week 12 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 37).Value
        Cells(i, 37).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 37).Value
    End If
Next i

' Week 13 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 38).Value
        Cells(i, 38).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 38).Value
    End If
Next i

' Week 14 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 39).Value
        Cells(i, 39).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 39).Value
    End If
Next i

' Week 15 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 40).Value
        Cells(i, 40).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 40).Value
    End If
Next i

' Week 16 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 46).Value
        Cells(i, 46).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 46).Value
    End If
Next i

' Week 17 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 47).Value
        Cells(i, 47).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 47).Value
    End If
Next i

' Week 18 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 48).Value
        Cells(i, 48).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 48).Value
    End If
Next i

' Week 19 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 49).Value
        Cells(i, 49).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 49).Value
    End If
Next i

' Week 20 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 50).Value
        Cells(i, 50).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 50).Value
    End If
Next i

' Week 21 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 56).Value
        Cells(i, 56).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 56).Value
    End If
Next i

' Week 22 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 57).Value
        Cells(i, 57).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 57).Value
    End If
Next i

' Week 23 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 58).Value
        Cells(i, 58).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 58).Value
    End If
Next i

' Week 24 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 59).Value
        Cells(i, 59).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 59).Value
    End If
Next i

' Week 25 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 60).Value
        Cells(i, 60).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 60).Value
    End If
Next i

' Week 26 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 66).Value
        Cells(i, 66).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 66).Value
    End If
Next i

' Week 27 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 67).Value
        Cells(i, 67).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 67).Value
    End If
Next i

' Week 28 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 68).Value
        Cells(i, 68).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 68).Value
    End If
Next i

' Week 29 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 69).Value
        Cells(i, 69).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 69).Value
    End If
Next i

' Week 30 total
open_qty_total = 0
For i = 2 To 11000
    If Cells(i, 2).Value <> Cells(i + 1, 2).Value Then
        open_qty_total = open_qty_total + Cells(i, 70).Value
        Cells(i, 70).Value = open_qty_total
        open_qty_total = 0
    Else
        open_qty_total = open_qty_total + Cells(i, 70).Value
    End If
Next i

' Page Layout '

    Cells.Select
    Selection.Font.Size = 10
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 80
    Range("A1").Select

' Formatting Row 1 '

    Range("A1:CM1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

' Labeling and formating first row

    Range("C1").Select
    ActiveCell.FormulaR1C1 = "C"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Primary Vendor"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Style-Color"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "*"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Beg Bal"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Ship"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Open"
    Range("M1").Select

' General Formatting

    Columns("H:H").Select
    Selection.InsertIndent 1
    Selection.ColumnWidth = 11
    Columns("I:I").Select
    Selection.ColumnWidth = 3
    Columns("I:CK").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "mo1"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(""Frcst "",MONTH(RC[6]),""/"",YEAR(RC[6]))"
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "PO"
    Range("P1:T1").Select
    Selection.NumberFormat = "[$-en-US]d-mmm;@"
    Range("U1").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(""Rcpts "",MONTH(RC[-1]),""/"",YEAR(RC[-1]))"
    Range("V1").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(""Bal "",MONTH(RC[-2]),""/"",YEAR(RC[-2]))"
    Range("W1").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(""% of "",MONTH(RC[-3]),""/"",YEAR(RC[-3]))"
    Range("X1").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(""Frcst "",MONTH(RC[6]),""/"",YEAR(RC[6]))"
    Range("Y1").Select
    ActiveCell.FormulaR1C1 = "PO"
    Range("U1:Y1").Select
    Selection.Copy
    Range("AE1").Select
    ActiveSheet.Paste
    Range("AO1").Select
    ActiveSheet.Paste
    Range("AY1").Select
    ActiveSheet.Paste
    Range("BI1").Select
    ActiveSheet.Paste
    Range("BS1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=CONCATENATE(""Rcpts "",MONTH(RC[-1]),""/"",YEAR(RC[-1]))"
    Range("BT1").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(""Bal "",MONTH(RC[-2]),""/"",YEAR(RC[-2]))"
    Range("BU1").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(""% of "",MONTH(RC[-3]),""/"",YEAR(RC[-3]))"
    Range("BV1").Select
    ActiveCell.FormulaR1C1 = "Frcst"
    Range("BW1").Select
    ActiveCell.FormulaR1C1 = "Frcst"
    Range("BX1").Select
    Range("Z1:AD1").Select
    Selection.NumberFormat = "[$-en-US]d-mmm;@"
    Range("AJ1:AN1").Select
    Selection.NumberFormat = "[$-en-US]d-mmm;@"
    Range("AT1:AX1").Select
    Selection.NumberFormat = "[$-en-US]d-mmm;@"
    Range("BD1:BH1").Select
    Selection.NumberFormat = "[$-en-US]d-mmm;@"
    Range("BN1:BR1").Select
    Selection.NumberFormat = "[$-en-US]d-mmm;@"
    Columns("W:W").Select
    Selection.Style = "Percent"
    Columns("AG:AG").Select
    Selection.Style = "Percent"
    Columns("AQ:AQ").Select
    Selection.Style = "Percent"
    Columns("BA:BA").Select
    Selection.Style = "Percent"
    Columns("BK:BK").Select
    Selection.Style = "Percent"
    Columns("BU:BU").Select
    Selection.Style = "Percent"
    Range("BX1").Select
    ActiveCell.FormulaR1C1 = "SH"
    Range("BY1").Select
    ActiveCell.FormulaR1C1 = "PO"
    Range("BZ1").Select
    ActiveCell.FormulaR1C1 = "S"
    Range("CB1").Select
    ActiveCell.FormulaR1C1 = "SC"
    Range("CE1").Select
    ActiveCell.FormulaR1C1 = "Orig Ship"
    Range("CF1").Select
    ActiveCell.FormulaR1C1 = "Req Ship"
    Range("CG1").Select
    ActiveCell.FormulaR1C1 = "Ship date"
    Range("CI1").Select
    ActiveCell.FormulaR1C1 = "ETA"
    Range("CJ1").Select
    ActiveCell.FormulaR1C1 = "Via"
    Range("CK1").Select
    ActiveCell.FormulaR1C1 = "DC"
    Columns("CE:CI").Select
    Selection.NumberFormat = "m/d;@"
    Selection.ColumnWidth = 5
    Columns("CM:CM").EntireColumn.AutoFit
    Columns("CJ:CK").Select
    Selection.ColumnWidth = 3
    Columns("CC:CC").Select
    Selection.ColumnWidth = 6
    Columns("CA:CB").Select
    Selection.ColumnWidth = 3
    Columns("BZ:BZ").Select
    Selection.ColumnWidth = 2
    Columns("BX:BX").Select
    Selection.ColumnWidth = 3
    Columns("J:BW").Select
    Selection.ColumnWidth = 6
    Columns("C:C").Select
    Selection.ColumnWidth = 1.5
    Range("S219").Select
    Columns("M:M").Select
    Selection.EntireColumn.Hidden = True
    Range("Q12").Select
    Range("AB28").Select
    Columns("BV:BW").Select
    Selection.EntireColumn.Hidden = True
    Columns("D:D").Select
    Selection.EntireColumn.Hidden = True
    Columns("F:G").Select
    Selection.Columns.Group
    Columns("O:T").Select
    Selection.Columns.Group
    Columns("N:V").Select
    Selection.Columns.Group
    Columns("Y:AD").Select
    Selection.Columns.Group
    Columns("X:AF").Select
    Selection.Columns.Group
    Columns("AI:AN").Select
    Selection.Columns.Group
    Columns("AH:AP").Select
    Selection.Columns.Group
    Columns("AS:AX").Select
    Selection.Columns.Group
    Columns("AR:AZ").Select
    Selection.Columns.Group
    Columns("BC:BH").Select
    Selection.Columns.Group
    Columns("BB:BJ").Select
    Selection.Columns.Group
    Columns("BM:BR").Select
    Selection.Columns.Group
    Columns("BL:BT").Select
    Selection.Columns.Group
    Range("BR5").Select
    Columns("BZ:CB").Select
    Selection.Columns.Group
    Columns("CF:CH").Select
    Selection.Columns.Group
    Columns("CM:CM").Select
    Selection.Columns.Group

' ETA Column

    Columns("CI:CI").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$CG1>0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

' PO column

    Columns("BY:BY").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="N", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

' Seq Cut Number column

    Columns("BZ:BZ").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False