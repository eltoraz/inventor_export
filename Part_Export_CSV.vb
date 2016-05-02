'Note: need to test whether GoExcel object works with CSV files
'If not, can gather the data into a comma-delimited string and export manually
GoExcel.Open("I:\Cadd\_iLogic\Export\Part_Level.csv")

Dim Description, PartType

'Write values in order expected by DMT
GoExcel.CellValue("A2") = "BBN"     'A2: Company name (constant)
GoExcel.CellValue("B2") = iProperties.Value("Project", "Part Number")

Description = iProperties.Value("Project", "Description")
GoExcel.CellValue("C2") = Left(Description, 8)  'Search word, first 8 characters of description
GoExcel.CellValue("D2") = Description

GoExcel.CellValue("E2") = iProperties.Value("Custom", "ClassID")
GoExcel.CellValue("F2") = iProperties.Value("Custom", "UOM")
GoExcel.CellValue("G2") = iProperties.Value("Custom", "UOM")

PartType = iProperties.Value("Custom", "Type")
GoExcel.CellValue("H2") = PartType

'I2: Price per grouping (currently: "E", but will this always be the case?)
GoExcel.CellValue("I2") = "E"
GoExcel.CellValue("J2") = iProperties.Value("Custom", "Group")
GoExcel.CellValue("K2") = iProperties.Value("Custom", "Mfg.Comments")
GoExcel.CellValue("L2") = iProperties.Value("Custom", "Purchase Comments")
GoExcel.CellValue("M2") = iProperties.Value("Custom", "UOM")
GoExcel.CellValue("N2") = iProperties.Value("Custom", "Use Part Rev")
GoExcel.CellValue("O2") = iProperties.Value("Custom", "Serial Number")
GoExcel.CellValue("P2") = iProperties.Value("Custom", "Base Number")

GoExcel.CellValue("Q2") = "BBN"     'UOMClassID

'R2/S2/T2: Serial Mask/Mask example/Net Weight UOM:
'possibly only needed for manufactured parts (type "M")?
If StrComp(PartType, "M") Then
    GoExcel.CellValue("R2") = "NF"
    GoExcel.CellValue("S2") = "NF9999999"
    GoExcel.CellValue("T2") = "LB"
End If

GoExcel.Save

MessageBox.Show("iProperties successfully copied!")
