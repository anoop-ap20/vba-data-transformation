Sub Project2()
Workbooks.Open Filename:="'Input file path - raw ERP purchase data xlsx file. named: Input.xlsx'" 'Like "D:\anoop-ap20\vba-data-transformation\Input.xlsx"
Workbooks.Add.SaveAs Filename:="'Output file path - the ingestion template with data'"
Workbooks.Open Filename:="'Base file path - the template you need. named: Based.xlsx'"

' Finding the last row in the Sheet
Dim a As Long, i As Long
a = Workbooks("Input.xlsx").Sheets(1).Range("A1048576").End(xlUp).Row
For x = 3 To a
If Workbooks("Input.xlsx").Sheets(1).Cells(x, 1).Value = "" Then
i = x - 1
Exit For
End If
Next x
Workbooks("Base.xlsx").Sheets(1).Range("A1:BI3").Copy
Workbooks("Output.xlsx").Sheets(1).Range("A1:BI3").PasteSpecial
Application.CutCopyMode = False

' B2B Invoices
'Invoice Date
Workbooks("Input.xlsx").Sheets(1).Range("D4:D" & i).Copy
Workbooks("Output.xlsx").Sheets(1).Range("A4").PasteSpecial
Application.CutCopyMode = False

'Invoice Number
Workbooks("Input.xlsx").Sheets(1).Range("C4:C" & i).Copy
Workbooks("Output.xlsx").Sheets(1).Range("B4").PasteSpecial
Application.CutCopyMode = False

'Supplier Name
Workbooks("Input.xlsx").Sheets(1).Range("B4:B" & i).Copy
Workbooks("Output.xlsx").Sheets(1).Range("C4").PasteSpecial
Application.CutCopyMode = False

'Supplier GSTIN
Workbooks("Input.xlsx").Sheets(1).Range("A4:A" & i).Copy
Workbooks("Output.xlsx").Sheets(1).Range("D4").PasteSpecial
Application.CutCopyMode = False

'Taxable Value
Workbooks("Input.xlsx").Sheets(1).Range("G4:G" & i).Copy
Workbooks("Output.xlsx").Sheets(1).Range("M4").PasteSpecial
Application.CutCopyMode = False

'CGST Amount
Workbooks("Input.xlsx").Sheets(1).Range("I4:I" & i).Copy
Workbooks("Output.xlsx").Sheets(1).Range("O4").PasteSpecial
Application.CutCopyMode = False

'SGST Amount
Workbooks("Input.xlsx").Sheets(1).Range("J4:J" & i).Copy
Workbooks("Output.xlsx").Sheets(1).Range("Q4").PasteSpecial
Application.CutCopyMode = False

'IGST Amount
Workbooks("Input.xlsx").Sheets(1).Range("H4:H" & i).Copy
Workbooks("Output.xlsx").Sheets(1).Range("S4").PasteSpecial
Application.CutCopyMode = False

'CESS Amount
Workbooks("Input.xlsx").Sheets(1).Range("K4:K" & i).Copy
Workbooks("Output.xlsx").Sheets(1).Range("U4").PasteSpecial
Application.CutCopyMode = False

'ITC Claim Type
For x = 4 To i
If Workbooks("Input.xlsx").Sheets(1).Cells(x, 13).Value = "Input Service" Then
Workbooks("Output.xlsx").Sheets(1).Cells(x, 22).Value = "Input Service"
ElseIf Workbooks("Input.xlsx").Sheets(1).Cells(x, 13).Value = "Ineligible For Input" Then
Workbooks("Output.xlsx").Sheets(1).Cells(x, 22).Value = "Ineligible"
End If
Next x

'POS
Workbooks("Input.xlsx").Sheets(1).Range("L4:L" & i).Copy
Workbooks("Output.xlsx").Sheets(1).Range("AQ4").PasteSpecial
Application.CutCopyMode = False

'Rectification in Output file
Dim e As Long
e = Workbooks("Output.xlsx").Sheets(1).Range("A1048576").End(xlUp).Row
For x = 4 To e
If Workbooks("Output.xlsx").Sheets(1).Cells(x, 13).Value < 0 Then
Workbooks("Output.xlsx").Sheets(1).Cells(x, 13).EntireRow.Delete
x = x - 1
End If
Next x
i = i + 1
For x = i To a
If Workbooks("Input.xlsx").Sheets(1).Cells(x, 1).Value = "6C. Amendments to details" Then
i = x + 3
Exit For
End If
Next x
Dim j As Long, b As Long, c As Long
For x = i To a
If Workbooks("Input.xlsx").Sheets(1).Cells(x, 1).Value = "" Then
j = x - 1
Exit For
End If
Next x
b = Workbooks("Output.xlsx").Sheets(1).Range("A1048576").End(xlUp).Row
b = b + 1

'CDNs
'Credit/Debit Note Date *
Workbooks("Input.xlsx").Sheets(1).Activate
Workbooks("Input.xlsx").Sheets(1).Range(Cells(i, 7), Cells(j, 7)).Copy
Workbooks("Output.xlsx").Sheets(1).Range("AA" & b).PasteSpecial
Application.CutCopyMode = False

'Credit(C)/ Debit(D) Note Type *
c = Workbooks("ClearPR.xlsx").Sheets(1).Range("AA1048576").End(xlUp).Row
Workbooks("Output.xlsx").Sheets(1).Activate
Workbooks("Output.xlsx").Sheets(1).Range(Cells(b, 29), Cells(c, 29)).Value = "C"

'Credit/Debit Note Number *
Workbooks("Input.xlsx").Sheets(1).Activate
Workbooks("Input.xlsx").Sheets(1).Range(Cells(i, 6), Cells(j, 6)).Copy
Workbooks("Output.xlsx").Sheets(1).Range("AB" & b).PasteSpecial
Application.CutCopyMode = False

'Invoice Date *
Workbooks("Input.xlsx").Sheets(1).Range(Cells(i, 4), Cells(j, 4)).Copy
Workbooks("Output.xlsx").Sheets(1).Range("A" & b).PasteSpecial
Application.CutCopyMode = False

'Invoice Number *
Workbooks("Input.xlsx").Sheets(1).Range(Cells(i, 3), Cells(j, 3)).Copy
Workbooks("Output.xlsx").Sheets(1).Range("B" & b).PasteSpecial
Application.CutCopyMode = False

'Supplier Name
Workbooks("Input.xlsx").Sheets(1).Range(Cells(i, 2), Cells(j, 2)).Copy
Workbooks("Output.xlsx").Sheets(1).Range("C" & b).PasteSpecial
Application.CutCopyMode = False

'Supplier GSTIN
Workbooks("Input.xlsx").Sheets(1).Range(Cells(i, 5), Cells(j, 5)).Copy
Workbooks("Output.xlsx").Sheets(1).Range("D" & b).PasteSpecial
Application.CutCopyMode = False

'Item Taxable Value *
Workbooks("Input.xlsx").Sheets(1).Range(Cells(i, 10), Cells(j, 10)).Copy
Workbooks("Output.xlsx").Sheets(1).Range("M" & b).PasteSpecial
Application.CutCopyMode = False

'CGST Amount
Workbooks("Input.xlsx").Sheets(1).Range(Cells(i, 12), Cells(j, 12)).Copy
Workbooks("Output.xlsx").Sheets(1).Range("O" & b).PasteSpecial
Application.CutCopyMode = False

'SGST Amount
Workbooks("Input.xlsx").Sheets(1).Range(Cells(i, 13), Cells(j, 13)).Copy
Workbooks("Output.xlsx").Sheets(1).Range("Q" & b).PasteSpecial
Application.CutCopyMode = False

'IGST Amount
Workbooks("Input.xlsx").Sheets(1).Range(Cells(i, 11), Cells(j, 11)).Copy
Workbooks("Output.xlsx").Sheets(1).Range("S" & b).PasteSpecial
Application.CutCopyMode = False

'CESS Amount
Workbooks("Input.xlsx").Sheets(1).Range(Cells(i, 14), Cells(j, 14)).Copy
Workbooks("Output.xlsx").Sheets(1).Range("U" & b).PasteSpecial
Application.CutCopyMode = False
'POS
Workbooks("Input.xlsx").Sheets(1).Range(Cells(i, 15), Cells(j, 15)).Copy
Workbooks("Output.xlsx").Sheets(1).Range("AQ" & b).PasteSpecial
Application.CutCopyMode = False

'ITC Claim Type
Dim d As Long
d = b
Workbooks("Input.xlsx").Sheets(1).Activate
For x = i To j
If Workbooks("Input.xlsx").Sheets(1).Cells(x, 16).Value = "Input Service" Then
Workbooks("Output.xlsx").Sheets(1).Cells(d, 22).Value = "Input Service"
d = d + 1
ElseIf Workbooks("Input.xlsx").Sheets(1).Cells(x, 16).Value = "Ineligible For Input" Then
Workbooks("Output.xlsx").Sheets(1).Cells(d, 22).Value = "Ineligible"
d = d + 1
Else
d = d + 1
End If
Next x
For x = 4 To c
If Workbooks("Output.xlsx").Sheets(1).Cells(x, 4).Value = "- None -" Then
Workbooks("Output.xlsx").Sheets(1).Cells(x, 4).EntireRow.Delete
x = x - 1
End If
Next x

'Correction of Invoice no.
Workbooks("Output.xlsx").Sheets(1).Range("B4:B" & c).Replace What:="Bill #",
Replacement:=""
Workbooks("Output.xlsx").Sheets(1).Range("B4:B" & c).Replace What:="Bill Credit #",
Replacement:=""
Workbooks("Output.xlsx").Sheets(1).Range("AB4:AB" & c).Replace What:="Bill Credit #",
Replacement:=""
Workbooks("Output.xlsx").Sheets(1).Range("AB4:AB" & c).Replace What:="Bill #",
Replacement:=""

Workbooks("Output.xlsx").Save
Workbooks("Output.xlsx").Close
Workbooks("Input.xlsx").Save
Workbooks("Input.xlsx").Close
Workbooks("Base.xlsx").Save
Workbooks("Base.xlsx").Close
End Sub
