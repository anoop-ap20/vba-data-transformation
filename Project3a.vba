Sub BifrucatingtheInputFile()
Workbooks.Open
Filename:="'Input file path - raw ERP purchase data xlsx file. named: Input.xlsx'" 'Like "D:\anoop-ap20\vba-data-transformation\Input.xlsx

‘ Eligible or Ineligible - ITC Claim Type
Dim ws As Worksheet
Dim lastRow As Long
Dim dict As Object
Dim cell As Range
Dim value As Variant
Set ws = Workbooks("Input.xlsx").Sheets(1)
lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
Set dict = CreateObject("Scripting.Dictionary")
For Each cell In ws.Range("E1:E" & lastRow)
value = cell.value
If dict.exists(value) Then
dict(value) = dict(value) + cell.Offset(0, 12).value
Else
dict.Add value, cell.Offset(0, 12).value
End If
Next cell
For Each cell In ws.Range("E1:E" & lastRow)
value = cell.value
If dict(value) = 0 Then
cell.Offset(0, 19).value = "INELIGIBLE"
Else
cell.Offset(0, 19).value = " "
End If
Next cell
With Workbooks("Input.xlsx")
Rows("1:1").Select
Selection.AutoFilter

'Cancelled Invoices
ActiveSheet.Range("$A$1:$AL$10790").AutoFilter Field:=31, Criteria1:= "*CANCELL*", Operator:=xlFilterValues
On Error Resume Next
ActiveSheet.Range("AE1:AE" & Workbooks(“Input.xlsx”).Sheets(1).Cells(Workbooks(“Input.xlsx”).Sheets(1).Workbooks(“Input.xlsx”).Sheets(1).Rows.Count,"AE").End(xlUp).Row).Offset(1,0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
On Error GoTo 0

'Import1
ActiveSheet.Range("$A$1:$AL$10790").AutoFilter Field:=32, Criteria1:= _
"<>" & "REVERSE CHARGE MECHANISIM"
ActiveSheet.Range("$A$1:$AL$10790").AutoFilter Field:=12, Criteria1:= _
"NOT APPLICABLE"
Cells.Select
Selection.Copy
Sheets.Add After:=ActiveSheet
Sheets("Sheet2").Select
Sheets("Sheet2").Name = "Import1"
ActiveSheet.Paste
Sheets("Sheet1").Select
Application.CutCopyMode = False
ActiveSheet.ShowAllData

'Import2
ActiveSheet.Range("$A$1:$AL$10790").AutoFilter Field:=7, Criteria1:= _
"PAYABLE"
Cells.Select
Selection.Copy
Sheets.Add After:=ActiveSheet
Sheets("Sheet3").Select
Sheets("Sheet3").Name = "Import2"
ActiveSheet.Paste
Sheets("Sheet1").Select
Application.CutCopyMode = False
ActiveSheet.ShowAllData

'RCM
ActiveSheet.Range("$A$1:$AL$10790").AutoFilter Field:=32, Criteria1:= _
"REVERSE CHARGE MECHANISIM", Operator:=xlOr, Criteria2:= _
"SERVICE INV WO PO-RCM"
Cells.Select
Selection.Copy
Sheets.Add After:=ActiveSheet
Sheets("Sheet4").Select
Sheets("Sheet4").Name = "RCM"
ActiveSheet.Paste
Sheets("Sheet1").Select
Application.CutCopyMode = False
ActiveSheet.ShowAllData

'Debit
ActiveSheet.Range("$A$1:$AL$10790").AutoFilter Field:=32, Criteria1:= _
"DEBIT"
Cells.Select
Selection.Copy
Sheets.Add After:=ActiveSheet
Sheets("Sheet5").Select
Sheets("Sheet5").Name = "Debit"
ActiveSheet.Paste
Sheets("Sheet1").Select
Application.CutCopyMode = False
ActiveSheet.ShowAllData

'Credit
ActiveSheet.Range("$A$1:$AL$10790").AutoFilter Field:=32, Criteria1:= _
"*CREDIT*"
Cells.Select
Selection.Copy
Sheets.Add After:=ActiveSheet
Sheets("Sheet6").Select
Sheets("Sheet6").Name = "Credit"
ActiveSheet.Paste
Sheets("Sheet1").Select
Application.CutCopyMode = False
ActiveSheet.ShowAllData

'Rest
ActiveSheet.Range("$A$1:$AL$10790").AutoFilter Field:=32, Criteria1:=”<>” & "DEBIT"
ActiveSheet.Range("$A$1:$AL$10790").AutoFilter Field:=32, Criteria1:="<>" & "*REVERSE CHARGE MECHANISIM*", Operator:=xlAnd, Criteria2:="<>SERVICE INV WO PO-RCM"
ActiveSheet.Range("$A$1:$AL$10790").AutoFilter Field:=32, Criteria1:="<>" & "*CREDIT*"
ActiveSheet.Range("$A$1:$AL$10790").AutoFilter Field:=12, Criteria1:= _
"<>" & "NOT APPLICABLE"
ActiveSheet.Range("$A$1:$AL$10790").AutoFilter Field:=7, Criteria1:= _
"<>" & "PAYABLE"
Cells.Select
Selection.Copy
Sheets.Add After:=ActiveSheet
Sheets("Sheet7").Select
Sheets("Sheet7").Name = "Rest"
ActiveSheet.Paste
Sheets("Sheet1").Select
Application.CutCopyMode = False
ActiveSheet.ShowAllData
End With

Workbooks("Input.xlsx").Save
Workbooks("Input.xlsx").Close
End Sub
