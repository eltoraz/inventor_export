Dim fso, FileName, csv
Dim Fields, Data
Dim Description, PartType, UOM

'Open the CSV file (note: this will overwrite the file if it exists!)
fso = CreateObject("Scripting.FileSystemObject")
FileName = "I:\Cadd\_iLogic\Export\Part_Level.csv"
csv = fso.OpenTextFile(FileName, 2, True, -2)

Fields = "Company,PartNum,SearchWord,PartDescription,ClassID,IUM,PUM,TypeCode,PricePerCode,ProdCode,MfgComment,PurComment,SalesUM,UsePartRev,SNFormat,SNBaseDataType,UOMClassID,SNMask,SNMaskExample,NetWeightUOM"

'Build string containing values in order expected by DMT (see Fields string)
Data = "BBN"                                'Company name (constant)
Data = Data & "," & iProperties.Value("Project", "Part Number")

Description = iProperties.Value("Project", "Description")
Data = Data & "," & Left(Description, 8)    'Search word, first 8 characters of description
Data = Data & "," & Description

Data = Data & "," & iProperties.Value("Custom", "ClassID")

if StrComp(PartType, "M") = 0 Then
    UOM = "EAM"
ElseIf StrComp(PartType, "P") = 0 Then
    UOM = "EAP"
End If

Data = Data & "," & UOM
Data = Data & "," & UOM

PartType = iProperties.Value("Custom", "Type")
Data = Data & "," & PartType

'Price per grouping (currently: "E", but will this always be the case?)
Data = Data & "," & "E"

Data = Data & "," & iProperties.Value("Custom", "Group")
Data = Data & "," & iProperties.Value("Custom", "Mfg.Comments")
Data = Data & "," & iProperties.Value("Custom", "Purchase Comments")
Data = Data & "," & UOM
Data = Data & "," & iProperties.Value("Custom", "Use Part Rev")
Data = Data & "," & iProperties.Value("Custom", "Serial Number")
Data = Data & "," & iProperties.Value("Custom", "Base Number")

Data = Data & "," & "BBN"                   'UOMClassID

'Serial Mask/Mask example/Net Weight UOM:
'possibly only needed for manufactured parts (type "M")?
If StrComp(PartType, "M") = 0 Then
    Data = Data & "," & "NF"
    Data = Data & "," & "NF9999999"
    Data = Data & "," & "LB"
End If

'Write field headers & data to file
csv.WriteLine(Fields)
csv.WriteLine(Data)
csv.Close()

MessageBox.Show("iProperties successfully copied!")
