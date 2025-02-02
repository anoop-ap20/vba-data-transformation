Sub Project1()
Workbooks.Open Filename:="'Input file path - raw ERP purchase data xlsx file. named: Input.xlsx'" 'Like "D:\anoop-ap20\vba-data-transformation\Input.xlsx"
Range("A:AA").Sort key1:=Range("B:B"), order1:=xlDescending, Header:=xlYes
Workbooks.Open Filename:="'Base file path - the template you need. named: Based.xlsx'"
Workbooks.Add.SaveAs Filename:="'Output file path - the ingestion template with data'"
Workbooks("Base.xlsx").Sheets(1).Unprotect Password:=123
Workbooks("Base.xlsx").Sheets(2).Unprotect Password:=123

'Finding the last cell
Dim a As Long
a = Workbooks("Input.xlsx").Sheets(1).Range("A1048576").End(xlUp).Row

'Ingestion template duplicated
Workbooks("Base.xlsx").Sheets(1).Range("A1:EM2").Copy
Workbooks("Output.xlsx").Sheets(1).Range("A1:EM2").PasteSpecial
Application.CutCopyMode = False

'Custom Logic
Workbooks("Output.xlsx").Sheets(1).Cells(2, 144).Value = "Other Details"
Workbooks("Output.xlsx").Sheets(1).Cells(2, 145).Value = "Job Ref."

'Inv Date and no.
Workbooks("Input.xlsx").Sheets(1).Activate
Workbooks("Input.xlsx").Sheets(1).Range("A2", Cells(a, 2)).Copy
Workbooks("Output.xlsx").Sheets(1).Activate
Workbooks("Output.xlsx").Sheets(1).Range("A3", Cells(a + 1, 2)).PasteSpecial
Application.CutCopyMode = False

'Customer Name
Workbooks("Input.xlsx").Sheets(1).Activate
Workbooks("Input.xlsx").Sheets(1).Range("C2", Cells(a, 3)).Copy
Workbooks("Output.xlsx").Sheets(1).Activate
Workbooks("Output.xlsx").Sheets(1).Range("E3", Cells(a + 1, 5)).PasteSpecial
Application.CutCopyMode = False

'Customer GSTIN
For i = 2 To a
Workbooks("Input.xlsx").Sheets(1).Activate
If StrComp("IN", Right(Cells(i, 4), 2), vbTextCompare) = 0 Then
Workbooks("Output.xlsx").Sheets(1).Cells(i + 1, 7).Value =
Workbooks("Input.xlsx").Sheets(1).Cells(i, 5).Value
End If
Next i

'Place of supply
Workbooks("Input.xlsx").Sheets(1).Activate
Workbooks("Input.xlsx").Sheets(1).Range("F2", Cells(a, 6)).Copy
Workbooks("Output.xlsx").Sheets(1).Activate
Workbooks("Output.xlsx").Sheets(1).Range("H3", Cells(a + 1, 8)).PasteSpecial
Application.CutCopyMode = False

'Customer Address and Place
Workbooks("Input.xlsx").Sheets(1).Activate
Workbooks("Input.xlsx").Sheets(1).Range("D2", Cells(a, 4)).Copy
Workbooks("Output.xlsx").Sheets(1).Activate
Workbooks("Output.xlsx").Sheets(1).Range("I3", Cells(a + 1, 9)).PasteSpecial
Application.CutCopyMode = False

'Customer State
Workbooks("Input.xlsx").Sheets(1).Activate
Workbooks("Input.xlsx").Sheets(1).Range("G2", Cells(a, 7)).Copy
Workbooks("Output.xlsx").Sheets(1).Activate
Workbooks("Output.xlsx").Sheets(1).Range("K3", Cells(a + 1, 11)).PasteSpecial
Workbooks("Output.xlsx").Sheets(1).Range("J3", Cells(a + 1, 10)).PasteSpecial
Application.CutCopyMode = False

'Cutomer Pin code
For j = 2 To a
Workbooks("Input.xlsx").Sheets(1).Activate
If StrComp("IN", Right(Cells(j, 4), 2), vbTextCompare) = 0 Then
Workbooks("Output.xlsx").Sheets(1).Cells(j + 1, 12).Value =
Workbooks("Input.xlsx").Sheets(1).Cells(j, 8).Value
Else
Workbooks("Output.xlsx").Sheets(1).Cells(j + 1, 12).Value = 999999
End If
Next j

'Sl no.
Dim l As Integer
Workbooks("Output.xlsx").Sheets(1).Activate
Dim b As Long
b = Workbooks("Output.xlsx").Sheets(1).Range("A1048576").End(xlUp).Row
For j = 3 To b
l = 1
For k = j To b
If StrComp(Cells(k, 2).Value, Cells(j, 2).Value, vbcompare) = 0 Then
Cells(k, 13).Value = l
l = l + 1
Else
Exit For
End If
Next k
j = k - 1
Next j

'Item Discription
Workbooks("Input.xlsx").Sheets(1).Activate
Workbooks("Input.xlsx").Sheets(1).Range("O2", Cells(a, 15)).Copy
Workbooks("Output.xlsx").Sheets(1).Activate
Workbooks("Output.xlsx").Sheets(1).Range("N3", Cells(a + 1, 14)).PasteSpecial
Application.CutCopyMode = False

'HSN Code
Workbooks("Input.xlsx").Sheets(1).Activate
Workbooks("Input.xlsx").Sheets(1).Range("I2", Cells(a, 9)).Copy
Workbooks("Output.xlsx").Sheets(1).Activate
Workbooks("Output.xlsx").Sheets(1).Range("P3", Cells(a + 1, 16)).PasteSpecial
Application.CutCopyMode = False

'Goods/Services
Workbooks("Output.xlsx").Sheets(1).Activate
For x = 3 To b
If Left(Cells(x, 16), 2) = 99 Then
Cells(x, 15).Value = "S"
ElseIf Left(Cells(x, 16), 2) <> 99 And Cells(x, 16).Value <> "" Then
Cells(x, 15).Value = "G"
End If
Next x

'Item price, gross amount, Item taxable vlaue
Workbooks("Input.xlsx").Sheets(1).Activate
Workbooks("Input.xlsx").Sheets(1).Range("J2", Cells(a, 10)).Copy
Workbooks("Output.xlsx").Sheets(1).Activate
Workbooks("Output.xlsx").Sheets(1).Range("S3", Cells(a + 1, 19)).PasteSpecial
Workbooks("Output.xlsx").Sheets(1).Range("T3", Cells(a + 1, 20)).PasteSpecial
Workbooks("Output.xlsx").Sheets(1).Range("V3", Cells(a + 1, 22)).PasteSpecial
Application.CutCopyMode = False

'GST Rate
For y = 2 To a
Workbooks("Output.xlsx").Sheets(1).Cells(y + 1, 23).Value =
Round((Workbooks("Input.xlsx").Sheets(1).Cells(y, 11).Value +
Workbooks("Input.xlsx").Sheets(1).Cells(y, 12).Value +
Workbooks("Input.xlsx").Sheets(1).Cells(y, 13).Value) * 100 /
Workbooks("Input.xlsx").Sheets(1).Cells(y, 10).Value, 0)
Next y

'Item Total amount
For Z = 3 To b
Workbooks("Output.xlsx").Sheets(1).Cells(Z, 34).Value =
Workbooks("Output.xlsx").Sheets(1).Cells(Z, 22).Value +
(Workbooks("Output.xlsx").Sheets(1).Cells(Z, 22).Value *
Workbooks("Output.xlsx").Sheets(1).Cells(Z, 23).Value / 100)
Next Z

'Total Taxable
Workbooks("Output.xlsx").Sheets(1).Activate
Dim e As Double
For c = 3 To b
e = 0
For d = c To b
If StrComp(Cells(d, 2).Value, Cells(c, 2).Value, vbcompare) = 0 Then
e = e + Cells(d, 22).Value
Else
Exit For
End If
Next d
For g = c To b
If StrComp(Cells(g, 2).Value, Cells(c, 2).Value, vbcompare) = 0 Then
Cells(g, 35).Value = e
Else
Exit For
End If
Next g
c = g - 1
Next c

'Total Invoice Value
Workbooks("Output.xlsx").Sheets(1).Activate
Dim o As Double
For m = 3 To b
o = 0
For n = m To b
If StrComp(Cells(n, 2).Value, Cells(m, 2).Value, vbcompare) = 0 Then
o = o + Cells(n, 34).Value
Else
Exit For
End If
Next n
For p = m To b
If StrComp(Cells(p, 2).Value, Cells(m, 2).Value, vbcompare) = 0 Then
Cells(p, 42).Value = o
Else
Exit For
End If
Next p
m = p - 1
Next m
Workbooks("Output.xlsx").Sheets(1).Activate
For r = 3 To b
Cells(r, 44).Value = Round(Cells(r, 42).Value, 0)
Next r

'Round off amount
For q = 3 To b
Cells(q, 43).Value = Cells(q, 42).Value - Cells(q, 44).Value
Next q
Range("AP3", Cells(b, 42)).Clear

'Supplier details and tax scheme
Workbooks("Output.xlsx").Sheets(1).Activate
For s = 3 To b
Cells(s, 49).Value = "SPS Intermodal Services India Private Limited"
Cells(s, 50).Value = "27AANCS4956R1Z7"
Cells(s, 51).Value = "Land Survey No 35 36 108, At Dighode, Dighode, Raigad, Maharashtra,"
Cells(s, 52).Value = "Navi Mumbai"
Cells(s, 53).Value = 27
Cells(s, 54).Value = 410206
Cells(s, 78).Value = "GST"
Next s

'Document type code
For t = 3 To b
If Cells(t, 22).Value < 0 Then
Cells(t, 3).Value = "CRN"
Else
Cells(t, 3).Value = "INV"
End If
Next t

'Supply type code
For u = 3 To b
If Cells(u, 7).Value <> "" Then
Cells(u, 4).Value = "B2B"
Else
Cells(u, 4).Value = "B2C"
End If
Next u

'Custom fields
Workbooks("Input.xlsx").Sheets(1).Activate
Range("P2", Cells(a, 26)).Copy
Workbooks("Output.xlsx").Sheets(1).Activate
Range("EB3", Cells(a + 1, 142)).PasteSpecial
Application.CutCopyMode = False

'Value is USD
Dim ac As Long
Workbooks("Base.xlsx").Sheets(2).Activate
ac = Workbooks("Base.xlsx").Sheets(2).Cells(2, 2).Value
Workbooks("Output.xlsx").Sheets(1).Activate
For ab = 3 To b
Workbooks("Output.xlsx").Sheets(1).Cells(ab, 143).Value =
Workbooks("Output.xlsx").Sheets(1).Cells(ab, 44).Value / ac
Next ab

'Merging columns
Workbooks("Output.xlsx").Sheets(1).Activate
For m = 3 To a + 1
Workbooks("Output.xlsx").Sheets(1).Cells(m, 144).Value = "Last Cargo: " &
Workbooks("Output.xlsx").Sheets(1).Cells(m, 136).Value & " " & "Move Number: " &
Workbooks("Output.xlsx").Sheets(1).Cells(m, 139).Value & " " & "From Date: " &
Workbooks("Output.xlsx").Sheets(1).Cells(m, 140).Value & " " & "To Date: " &
Workbooks("Output.xlsx").Sheets(1).Cells(m, 141).Value
Next m

'Custom Logic - Change in Item Disc.
For m = 3 To a + 1
If Workbooks("Output.xlsx").Sheets(1).Cells(m, 14).Value = "M&R" Then
Workbooks("Output.xlsx").Sheets(1).Cells(m, 14).Value = "Tanks Repair Charges" & " " &
"Container No: " & Workbooks("Output.xlsx").Sheets(1).Cells(m, 135).Value
ElseIf Workbooks("Output.xlsx").Sheets(1).Cells(m, 14).Value = "CLEANING" Then
Workbooks("Output.xlsx").Sheets(1).Cells(m, 14).Value = "Tank Cleaning Charges" & " " &
"Container No: " & Workbooks("Output.xlsx").Sheets(1).Cells(m, 135).Value
ElseIf Workbooks("Output.xlsx").Sheets(1).Cells(m, 14).Value = "STORAGE" Then
Workbooks("Output.xlsx").Sheets(1).Cells(m, 14).Value = "Tank Storage Charges" & " " &
"Container No: " & Workbooks("Output.xlsx").Sheets(1).Cells(m, 135).Value
ElseIf Workbooks("Output.xlsx").Sheets(1).Cells(m, 14).Value = "MISC" Then
Workbooks("Output.xlsx").Sheets(1).Cells(m, 14).Value = "Miscellaneous" & " " & "Container
No: " & Workbooks("Output.xlsx").Sheets(1).Cells(m, 135).Value
ElseIf Workbooks("Output.xlsx").Sheets(1).Cells(m, 14).Value = "LIFT ON" Then
Workbooks("Output.xlsx").Sheets(1).Cells(m, 14).Value = "Tank Lift On Charges" & " " &
"Container No: " & Workbooks("Output.xlsx").Sheets(1).Cells(m, 135).Value
ElseIf Workbooks("Output.xlsx").Sheets(1).Cells(m, 14).Value = "LIFT OFF" Then
Workbooks("Output.xlsx").Sheets(1).Cells(m, 14).Value = "Tank Lift Off Charges" & " " &
"Container No: " & Workbooks("Output.xlsx").Sheets(1).Cells(m, 135).Value
ElseIf Workbooks("Output.xlsx").Sheets(1).Cells(m, 14).Value = "EIR IN" Then
Workbooks("Output.xlsx").Sheets(1).Cells(m, 14).Value = "EIR IN" & " " & "Container No: " &
Workbooks("Output.xlsx").Sheets(1).Cells(m, 135).Value
End If
Next m

'Merging column 2
For m = 3 To a + 1
If Workbooks("Output.xlsx").Sheets(1).Cells(m, 137).Value <> "" Then
Workbooks("Output.xlsx").Sheets(1).Cells(m, 145).Value = "MR Job ID: " &
Workbooks("Output.xlsx").Sheets(1).Cells(m, 137).Value & " " & "Work Order No: " &
Workbooks("Output.xlsx").Sheets(1).Cells(m, 138).Value
ElseIf Workbooks("Output.xlsx").Sheets(1).Cells(m, 142).Value <> "" Then
Workbooks("Output.xlsx").Sheets(1).Cells(m, 145).Value = "Charge Days: " &
Workbooks("Output.xlsx").Sheets(1).Cells(m, 142).Value
End If
Next m

Cells(1, 1).Select
Workbooks("Base.xlsx").Sheets(1).Protect Password:=123
Workbooks("Base.xlsx").Sheets(2).Protect Password:=123
Workbooks("Output.xlsx").Save
Workbooks("Input.xlsx").Save
Workbooks("Base.xlsx").Save
Workbooks("Output.xlsx").Close
Workbooks("Input.xlsx").Close
Workbooks("Base.xlsx").Close
End Sub
