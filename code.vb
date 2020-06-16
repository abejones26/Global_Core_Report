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