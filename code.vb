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