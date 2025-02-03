Sub Project3b()
Workbooks.Open Filename:="'Input file path - raw ERP purchase data xlsx file. named: Input.xlsx'" 'Like "D:\anoop-ap20\vba-data-transformation\Input.xlsx
Workbooks.Add.SaveAs Filename:="'Output file path - the ingestion template with data'"
Workbooks.Open Filename:="'Base file path - the template you need. named: Based.xlsx'"
Dim wsa As Workbook
Dim wsb As Workbook
Set wsa = Workbooks("Input.xlsx")
Set wsb = Workbooks("Output.xlsx")
Workbooks("Base.xlsx").Sheets(1).Range("A1:BK3").Copy
wsb.Sheets(1).Range("A1:BK3").PasteSpecial
Application.CutCopyMode = False

'<< Rest (B2B) >>
Dim a As Long
a = wsa.Sheets("Rest").Range("A1048576").End(xlUp).Row
  
'Invoice Date
wsa.Sheets("Rest").Range("F2:F" & a).Copy
wsb.Sheets(1).Range("A4").PasteSpecial
Application.CutCopyMode = False
  
'Invoice Number
wsa.Sheets("Rest").Range("E2:E" & a).Copy
wsb.Sheets(1).Range("B4").PasteSpecial
Application.CutCopyMode = False
  
'Supplier Name
wsa.Sheets("Rest").Range("G2:G" & a).Copy
wsb.Sheets(1).Range("C4").PasteSpecial
Application.CutCopyMode = False
  
'Supplier GSTIN
wsa.Sheets("Rest").Range("L2:L" & a).Copy
wsb.Sheets(1).Range("D4").PasteSpecial
Application.CutCopyMode = False
  
'HSN or SAC code
wsa.Sheets("Rest").Range("K2:K" & a).Copy
wsb.Sheets(1).Range("H4").PasteSpecial
Application.CutCopyMode = False
  
'Item Unit of Measurement
wsa.Sheets("Rest").Range("M2:M" & a).Copy
wsb.Sheets(1).Range("J4").PasteSpecial
Application.CutCopyMode = False
  
'Item Quantity
wsa.Sheets("Rest").Range("N2:N" & a).Copy
wsb.Sheets(1).Range("I4").PasteSpecial
Application.CutCopyMode = False
  
'Taxable Value
wsa.Sheets("Rest").Range("Q2:Q" & a).Copy
wsb.Sheets(1).Range("M4").PasteSpecial
Application.CutCopyMode = False
  
'CGST Amount
wsa.Sheets("Rest").Range("T2:T" & a).Copy
wsb.Sheets(1).Range("O4").PasteSpecial
Application.CutCopyMode = False
  
'SGST Amount
wsa.Sheets("Rest").Range("U2:U" & a).Copy
wsb.Sheets(1).Range("Q4").PasteSpecial
Application.CutCopyMode = False
  
'IGST Amount
wsa.Sheets("Rest").Range("S2:S" & a).Copy
wsb.Sheets(1).Range("S4").PasteSpecial
Application.CutCopyMode = False
  
'My GSTIN
wsa.Sheets("Rest").Range("AH2:AH" & a).Copy
wsb.Sheets(1).Range("AP4").PasteSpecial
Application.CutCopyMode = False
  
'State Place of Supply
wsa.Sheets("Rest").Range("AG2:AG" & a).Copy
wsb.Sheets(1).Range("AQ4").PasteSpecial
Application.CutCopyMode = False
  
'Total Transaction Value
wsa.Sheets("Rest").Range("Y2:Y" & a).Copy
wsb.Sheets(1).Range("BG4").PasteSpecial
Application.CutCopyMode = False
  
'ITC Claim Type
wsa.Sheets("Rest").Range("X2:X” & a).Copy
wsb.Sheets(1).Range(“V4”).PasteSpecial
Application.CutCopyMode = False
    
'<< Credit (CRNs) >>
Dim b As Long, c As Long
b = wsa.Sheets("Credit").Range("A1048576").End(xlUp).Row
c = wsb.Sheets(1).Range("A1048576").End(xlUp).Row

'Credit/Debit Note Date *
wsa.Sheets("Credit").Range("F2:F" & b).Copy
wsb.Sheets(1).Cells(c + 1, 27).PasteSpecial
Application.CutCopyMode = False

'Credit/Debit Note Number *
wsa.Sheets("Credit").Range("E2:E" & b).Copy
wsb.Sheets(1).Cells(c + 1, 28).PasteSpecial
Application.CutCopyMode = False

'Supplier Name
wsa.Sheets("Credit").Range("G2:G" & b).Copy
wsb.Sheets(1).Cells(c + 1, 3).PasteSpecial
Application.CutCopyMode = False
'Supplier GSTIN
wsa.Sheets("Credit").Range("L2:L" & b).Copy
wsb.Sheets(1).Cells(c + 1, 4).PasteSpecial
Application.CutCopyMode = False

'HSN or SAC code
wsa.Sheets("Credit").Range("K2:K" & b).Copy
wsb.Sheets(1).Cells(c + 1, 8).PasteSpecial
Application.CutCopyMode = False

'Item Unit of Measurement
wsa.Sheets("Credit").Range("M2:M" & b).Copy
wsb.Sheets(1).Cells(c + 1, 10).PasteSpecial
Application.CutCopyMode = False

'Item Quantity
wsa.Sheets("Credit").Range("N2:N" & b).Copy
wsb.Sheets(1).Cells(c + 1, 9).PasteSpecial
Application.CutCopyMode = False

'Taxable Value
wsa.Sheets("Credit").Range("Q2:Q" & b).Copy
wsb.Sheets(1).Cells(c + 1, 13).PasteSpecial
Application.CutCopyMode = False

'CGST Amount
wsa.Sheets("Credit").Range("T2:T" & b).Copy
wsb.Sheets(1).Cells(c + 1, 15).PasteSpecial
Application.CutCopyMode = False

'SGST Amount
wsa.Sheets("Credit").Range("U2:U" & b).Copy
wsb.Sheets(1).Cells(c + 1, 17).PasteSpecial
Application.CutCopyMode = False

'IGST Amount
wsa.Sheets("Credit").Range("S2:S" & b).Copy
wsb.Sheets(1).Cells(c + 1, 19).PasteSpecial
Application.CutCopyMode = False

'My GSTIN
wsa.Sheets("Credit").Range("AH2:AH" & b).Copy
wsb.Sheets(1).Cells(c + 1, 42).PasteSpecial
Application.CutCopyMode = False

'State Place of Supply
wsa.Sheets("Credit").Range("AG2:AG" & b).Copy
wsb.Sheets(1).Cells(c + 1, 43).PasteSpecial
Application.CutCopyMode = False

'Total Transaction Value
wsa.Sheets("Credit").Range("Y2:Y" & b).Copy
wsb.Sheets(1).Cells(c + 1, 59).PasteSpecial
Application.CutCopyMode = False

'ITC Claim Type
wsa.Sheets("Credit").Range("X2:X" & b).Copy
wsb.Sheets(1).Cells(c + 1, 22).PasteSpecial
Application.CutCopyMode = False

'<< Debit (DBNs) >>
Dim d As Long, e As Long
e = wsa.Sheets("Debit").Range("A1048576").End(xlUp).Row
d = wsb.Sheets(1).Range("M1048576").End(xlUp).Row

'Credit(C)/ Debit(D) Note Type *
wsb.Sheets(1).Range("AC" & c + 1 & ":" & "AC" & d).Value = "C"

'Credit/Debit Note Date *
wsa.Sheets("Debit").Range("F2:F" & e).Copy
wsb.Sheets(1).Cells(d + 1, 27).PasteSpecial
Application.CutCopyMode = False

'Credit/Debit Note Number *
wsa.Sheets("Debit").Range("E2:E" & e).Copy
wsb.Sheets(1).Cells(d + 1, 28).PasteSpecial
Application.CutCopyMode = False

'Supplier Name
wsa.Sheets("Debit").Range("G2:G" & e).Copy
wsb.Sheets(1).Cells(d + 1, 3).PasteSpecial
Application.CutCopyMode = False

'Supplier GSTIN
wsa.Sheets("Debit").Range("L2:L" & e).Copy
wsb.Sheets(1).Cells(d + 1, 4).PasteSpecial
Application.CutCopyMode = False

'HSN or SAC code
wsa.Sheets("Debit").Range("K2:K" & e).Copy
wsb.Sheets(1).Cells(d + 1, 8).PasteSpecial
Application.CutCopyMode = False

'Item Unit of Measurement
wsa.Sheets("Debit").Range("M2:M" & e).Copy
wsb.Sheets(1).Cells(d + 1, 10).PasteSpecial
Application.CutCopyMode = False

'Item Quantity
wsa.Sheets("Debit").Range("N2:N" & e).Copy
wsb.Sheets(1).Cells(d + 1, 9).PasteSpecial
Application.CutCopyMode = False

'Taxable Value
wsa.Sheets("Debit").Range("Q2:Q" & e).Copy
wsb.Sheets(1).Cells(d + 1, 13).PasteSpecial
Application.CutCopyMode = False

'CGST Amount
wsa.Sheets("Debit").Range("T2:T" & e).Copy
wsb.Sheets(1).Cells(d + 1, 15).PasteSpecial
Application.CutCopyMode = False

'SGST Amount
wsa.Sheets("Debit").Range("U2:U" & e).Copy
wsb.Sheets(1).Cells(d + 1, 17).PasteSpecial
Application.CutCopyMode = False

'IGST Amount
wsa.Sheets("Debit").Range("S2:S" & e).Copy
wsb.Sheets(1).Cells(d + 1, 19).PasteSpecial
Application.CutCopyMode = False

'My GSTIN
wsa.Sheets("Debit").Range("AH2:AH" & e).Copy
wsb.Sheets(1).Cells(d + 1, 42).PasteSpecial
Application.CutCopyMode = False

'State Place of Supply
wsa.Sheets("Debit").Range("AG2:AG" & e).Copy
wsb.Sheets(1).Cells(d + 1, 43).PasteSpecial
Application.CutCopyMode = False

'Total Transaction Value
wsa.Sheets("Debit").Range("Y2:Y" & e).Copy
wsb.Sheets(1).Cells(d + 1, 59).PasteSpecial
Application.CutCopyMode = False

'ITC Claim Type
wsa.Sheets("Debit").Range("X2:X" & e).Copy
wsb.Sheets(1).Cells(d + 1, 22).PasteSpecial
Application.CutCopyMode = False

'<< Reverse Charge Mechanism (RCMs) >>
Dim f As Long, g As Long
f = wsa.Sheets("RCM").Range("A1048576").End(xlUp).Row
g = wsb.Sheets(1).Range("M1048576").End(xlUp).Row

'Credit(C)/ Debit(D) Note Type *
wsb.Sheets(1).Range("AC" & d + 1 & ":" & "AC" & g).Value = "C"

'Invoice Date
wsa.Sheets("RCM").Range("F2:F" & f).Copy
wsb.Sheets(1).Cells(g + 1, 1).PasteSpecial
Application.CutCopyMode = False

'Invoice Number
wsa.Sheets("RCM").Range("E2:E" & f).Copy
wsb.Sheets(1).Cells(g + 1, 2).PasteSpecial
Application.CutCopyMode = False

'Supplier Name
wsa.Sheets("RCM").Range("G2:G" & f).Copy
wsb.Sheets(1).Cells(g + 1, 3).PasteSpecial
Application.CutCopyMode = False

'Supplier GSTIN
wsa.Sheets("RCM").Range("L2:L" & f).Copy
wsb.Sheets(1).Cells(g + 1, 4).PasteSpecial
Application.CutCopyMode = False

'HSN or SAC code
wsa.Sheets("RCM").Range("K2:K" & f).Copy
wsb.Sheets(1).Cells(g + 1, 8).PasteSpecial
Application.CutCopyMode = False

'Item Unit of Measurement
wsa.Sheets("RCM").Range("M2:M" & f).Copy
wsb.Sheets(1).Cells(g + 1, 10).PasteSpecial
Application.CutCopyMode = False

'Item Quantity
wsa.Sheets("RCM").Range("N2:N" & f).Copy
wsb.Sheets(1).Cells(g + 1, 9).PasteSpecial
Application.CutCopyMode = False

'Taxable Value
wsa.Sheets("RCM").Range("Q2:Q" & f).Copy
wsb.Sheets(1).Cells(g + 1, 13).PasteSpecial
Application.CutCopyMode = False

'CGST Amount
wsa.Sheets("RCM").Range("T2:T" & f).Copy
wsb.Sheets(1).Cells(g + 1, 15).PasteSpecial
Application.CutCopyMode = False

'SGST Amount
wsa.Sheets("RCM").Range("U2:U" & f).Copy
wsb.Sheets(1).Cells(g + 1, 17).PasteSpecial
Application.CutCopyMode = False

'IGST Amount
wsa.Sheets("RCM").Range("S2:S" & f).Copy
wsb.Sheets(1).Cells(g + 1, 19).PasteSpecial
Application.CutCopyMode = False

'My GSTIN
wsa.Sheets("RCM").Range("AH2:AH" & f).Copy
wsb.Sheets(1).Cells(g + 1, 42).PasteSpecial
Application.CutCopyMode = False

'State Place of Supply
wsa.Sheets("RCM").Range("AG2:AG" & f).Copy
wsb.Sheets(1).Cells(g + 1, 43).PasteSpecial
Application.CutCopyMode = False

'Total Transaction Value
wsa.Sheets("RCM").Range("Y2:Y" & f).Copy
wsb.Sheets(1).Cells(g + 1, 59).PasteSpecial
Application.CutCopyMode = False

'ITC Claim Type
wsa.Sheets("RCM").Range("X2:X" & f).Copy
wsb.Sheets(1).Cells(g + 1, 22).PasteSpecial
Application.CutCopyMode = False

'Is Reverse Charge Applicable?
Dim h As Long, i As Long
h = wsb.Sheets(1).Range("M1048576").End(xlUp).Row
wsb.Sheets(1).Range("AG" & g + 1 & ":" & "AG" & h).Value = "Y"
i = wsa.Sheets("Import1").Range("A1048576").End(xlUp).Row

'<< Imports1 >>
'Invoice Date
wsa.Sheets("Import1").Range("F2:F" & i).Copy
wsb.Sheets(1).Cells(h + 1, 1).PasteSpecial
Application.CutCopyMode = False

'Invoice Number
wsa.Sheets("Import1").Range("E2:E" & i).Copy
wsb.Sheets(1).Cells(h + 1, 2).PasteSpecial
Application.CutCopyMode = False

'Supplier Name
wsa.Sheets("Import1").Range("G2:G" & i).Copy
wsb.Sheets(1).Cells(h + 1, 3).PasteSpecial
Application.CutCopyMode = False

'Supplier GSTIN
wsa.Sheets("Import1").Range("L2:L" & i).Copy
wsb.Sheets(1).Cells(h + 1, 4).PasteSpecial
Application.CutCopyMode = False

'HSN or SAC code
wsa.Sheets("Import1").Range("K2:K" & i).Copy
wsb.Sheets(1).Cells(h + 1, 8).PasteSpecial
Application.CutCopyMode = False

'Item Unit of Measurement
wsa.Sheets("Import1").Range("M2:M" & i).Copy
wsb.Sheets(1).Cells(h + 1, 10).PasteSpecial
Application.CutCopyMode = False

'Item Quantity
wsa.Sheets("Import1").Range("N2:N" & i).Copy
wsb.Sheets(1).Cells(h + 1, 9).PasteSpecial
Application.CutCopyMode = False

'Taxable Value
wsa.Sheets("Import1").Range("Q2:Q" & i).Copy
wsb.Sheets(1).Cells(h + 1, 13).PasteSpecial
Application.CutCopyMode = False

'CGST Amount
wsa.Sheets("Import1").Range("T2:T" & i).Copy
wsb.Sheets(1).Cells(h + 1, 15).PasteSpecial
Application.CutCopyMode = False

'SGST Amount
wsa.Sheets("Import1").Range("U2:U" & i).Copy
wsb.Sheets(1).Cells(h + 1, 17).PasteSpecial
Application.CutCopyMode = False

'IGST Amount
wsa.Sheets("Import1").Range("S2:S" & i).Copy
wsb.Sheets(1).Cells(h + 1, 19).PasteSpecial
Application.CutCopyMode = False

'My GSTIN
wsa.Sheets("Import1").Range("AH2:AH" & i).Copy
wsb.Sheets(1).Cells(h + 1, 42).PasteSpecial
Application.CutCopyMode = False

'State Place of Supply
wsa.Sheets("Import1").Range("AG2:AG" & i).Copy
wsb.Sheets(1).Cells(h + 1, 43).PasteSpecial
Application.CutCopyMode = False

'Total Transaction Value
wsa.Sheets("Import1").Range("Y2:Y" & i).Copy
wsb.Sheets(1).Cells(h + 1, 59).PasteSpecial
Application.CutCopyMode = False

'Bill of Entry Date
wsa.Sheets("Import1").Range("F2:F" & i).Copy
wsb.Sheets(1).Cells(h + 1, 37).PasteSpecial
Application.CutCopyMode = False

'Bill of Entry Number
wsa.Sheets("Import1").Range("E2:E" & i).Copy
wsb.Sheets(1).Cells(h + 1, 36).PasteSpecial
Application.CutCopyMode = False

'ITC Claim Type
wsa.Sheets("Import1").Range("X2:X" & i).Copy
wsb.Sheets(1).Cells(h + 1, 22).PasteSpecial
Application.CutCopyMode = False

'Bill of Entry Port Code
j = wsb.Sheets(1).Range("M1048576").End(xlUp).Row
k = wsa.Sheets("Import2").Range("A1048576").End(xlUp).Row

'<< Import2 >>
'Invoice Date
wsa.Sheets("Import2").Range("F2:F" & k).Copy
wsb.Sheets(1).Cells(j + 1, 1).PasteSpecial
Application.CutCopyMode = False

'Invoice Number
wsa.Sheets("Import2").Range("E2:E" & k).Copy
wsb.Sheets(1).Cells(j + 1, 2).PasteSpecial
Application.CutCopyMode = False

'Supplier Name
wsa.Sheets("Import2").Range("G2:G" & k).Copy
wsb.Sheets(1).Cells(j + 1, 3).PasteSpecial
Application.CutCopyMode = False

'Supplier GSTIN
wsa.Sheets("Import2").Range("L2:L" & k).Copy
wsb.Sheets(1).Cells(j + 1, 4).PasteSpecial
Application.CutCopyMode = False

'HSN or SAC code
wsa.Sheets("Import2").Range("K2:K" & k).Copy
wsb.Sheets(1).Cells(j + 1, 8).PasteSpecial
Application.CutCopyMode = False

'Item Unit of Measurement
wsa.Sheets("Import2").Range("M2:M" & k).Copy
wsb.Sheets(1).Cells(j + 1, 10).PasteSpecial
Application.CutCopyMode = False

'Item Quantity
wsa.Sheets("Import2").Range("N2:N" & k).Copy
wsb.Sheets(1).Cells(j + 1, 9).PasteSpecial
Application.CutCopyMode = False

'Taxable Value
wsa.Sheets("Import2").Range("Q2:Q" & k).Copy
wsb.Sheets(1).Cells(j + 1, 13).PasteSpecial
Application.CutCopyMode = False

'CGST Amount
wsa.Sheets("Import2").Range("T2:T" & k).Copy
wsb.Sheets(1).Cells(j + 1, 15).PasteSpecial
Application.CutCopyMode = False

'SGST Amount
wsa.Sheets("Import2").Range("U2:U" & k).Copy
wsb.Sheets(1).Cells(j + 1, 17).PasteSpecial
Application.CutCopyMode = False

'IGST Amount
wsa.Sheets("Import2").Range("S2:S" & k).Copy
wsb.Sheets(1).Cells(j + 1, 19).PasteSpecial
Application.CutCopyMode = False

'My GSTIN
wsa.Sheets("Import2").Range("AH2:AH" & k).Copy
wsb.Sheets(1).Cells(j + 1, 42).PasteSpecial
Application.CutCopyMode = False

'State Place of Supply
wsa.Sheets("Import2").Range("AG2:AG" & k).Copy
wsb.Sheets(1).Cells(j + 1, 43).PasteSpecial
Application.CutCopyMode = False

'Total Transaction Value
wsa.Sheets("Import2").Range("Y2:Y" & k).Copy
wsb.Sheets(1).Cells(j + 1, 59).PasteSpecial
Application.CutCopyMode = False

'Bill of Entry Date
wsa.Sheets("Import2").Range("F2:F" & k).Copy
wsb.Sheets(1).Cells(j + 1, 37).PasteSpecial
Application.CutCopyMode = False

'Bill of Entry Number
wsa.Sheets("Import2").Range("E2:E" & k).Copy
wsb.Sheets(1).Cells(j + 1, 36).PasteSpecial
Application.CutCopyMode = False

'ITC Claim Type
wsa.Sheets("Import2").Range("X2:X" & k).Copy
wsb.Sheets(1).Cells(j + 1, 22).PasteSpecial
Application.CutCopyMode = False

'Bill of Entry Port Code
k = wsb.Sheets(1).Range("M1048576").End(xlUp).Row

'Port Code - Custom Logic
wsb.Sheets.Add.Name = "Port Code"
Workbooks("Base.xlsx").Sheets("Port code").Range("A2:E637").Copy
wsb.Sheets("Port Code").Range("A2:E637").PasteSpecial
Application.CutCopyMode = False
Dim l As Long
Dim sk As Long
sk = h + 1
For l = h + 1 To k
wsb.Sheets(1).Cells(l, "AI").Formula = "=XLOOKUP(AQ" & sk & ",'Port code'!E:E,'Port code'!B:B,""NA"",0)"
sk = sk + 1
Next l
'Range("F2").Value = WorksheetFunction.Xlookup(Range("E2"), Range("A2:A11"), Range("C2:C11"))
'wsb.Sheets(1).Range("AI" & h + 1 & ":" & "AI" & k).Value =
WorksheetFunction.XLookup(Range("AQ" & h + 1 & ":" & "AQ" & k), wsb.Sheets("Port code").Range("E:E"), wsb.Sheets("Port code").Range("B:B"))
'wsb.Sheets(1).Range("D4:D" & k).Replace What:="NOT APPLICABLE", Replacement:=""
'wsb.Sheets(1).Range("D4:D" & k).Replace What:="UNREGISTERED", Replacement:="URP"
'wsb.Sheets(1).Range("B4:B" & k).Replace What:=",", Replacement:=""
'wsb.Sheets(1).Range("B4:B" & k).Replace What:=".", Replacement:=""
'wsb.Sheets(1).Range("AB4:AB" & k).Replace What:=",", Replacement:=""
'wsb.Sheets(1).Range("AB4:AB" & k).Replace What:=".", Replacement:=""
wsb.Save
wsb.Close
wsa.Save
wsa.Close
Workbooks("Base.xlsx").Save
Workbooks("Base.xlsx").Close
End Sub
