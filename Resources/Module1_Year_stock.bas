Attribute VB_Name = "Module1"
Sub Multiple_year_stock()
' Macro3 Macro
   Columns("A:A").Select
   Selection.Copy
   Columns("I:I").Select
   ActiveSheet.Paste
   Application.CutCopyMode = False
   ActiveSheet.Range("$I$1:$I$70926").RemoveDuplicates Columns:=1, Header:= _
       xlNo
   Range("I1").Select
   ActiveCell.FormulaR1C1 = "Ticker"
   Range("J1").Select
   ActiveCell.FormulaR1C1 = "Total Sum"
   Range("J2").Select
   ActiveCell.FormulaR1C1 = "=SUMIF(R2C1:R70926C[-9],RC[-1],R2C7:R70926C7)"
   Range("J2").Select
   Selection.AutoFill Destination:=Range("J2:J290")
   Range("J2:J290").Select
End Sub

