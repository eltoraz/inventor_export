Dim fso, FileName, csv
Dim Fields, Data
Dim Description, PartType, UOM, TrackSerialNum
Dim SNFormat, SNBaseDataType, SNMask, SNMaskExample

Description = iProperties.Value("Project", "Description")
PartType = iProperties.Value("Custom", "PartType")

'UOM is set based on part type
'NOTE: default is "M" (manufactured), though only "P"/"M" are expected
If StrComp(PartType, "P") = 0 Then
    UOM = "EAP"
Else
    UOM = "EAM"
End If

TrackSerialNum = iProperties.Value("Custom", "TrackSerialNum")
'if serial number is being tracked, a bunch of fields are enabled
If TrackSerialNum Then
    SNFormat = "NF#######"
    SNBaseDataType = "MASK"
    SNMask = "NF"
    SNMaskExample = "NF9999999"
Else
    SNFormat = ""
    SNBaseDataType = ""
    SNMask = ""
    SNMaskExample = ""
End If

'Open the CSV file (note: this will overwrite the file if it exists!)
fso = CreateObject("Scripting.FileSystemObject")
FileName = "I:\Cadd\_iLogic\Export\Part_Level.csv"
csv = fso.OpenTextFile(FileName, 2, True, -2)

Fields = "Company,PartNum,SearchWord,PartDescription,ClassID,IUM,PUM,TypeCode,PricePerCode,ProdCode,MfgComment,PurComment,TrackSerialNum,SalesUM,UsePartRev,SNFormat,SNBaseDataType,UOMClassID,SNMask,SNMaskExample,NetWeightUOM"

'Build string containing values in order expected by DMT (see Fields string)
Data = "BBN"                                'Company name (constant)
Data = Data & "," & iProperties.Value("Project", "Part Number")

Data = Data & "," & Left(Description, 8)    'Search word, first 8 characters of description
Data = Data & "," & Description

Data = Data & "," & iProperties.Value("Custom", "ClassID")

Data = Data & "," & UOM
Data = Data & "," & UOM

Data = Data & "," & PartType

'Price per grouping (currently: "E", but will this always be the case?)
Data = Data & "," & "E"

Data = Data & "," & iProperties.Value("Custom", "ProdCode")
Data = Data & "," & iProperties.Value("Custom", "MfgComment")
Data = Data & "," & iProperties.Value("Custom", "PurComment")
Data = Data & "," & TrackSerialNum
Data = Data & "," & UOM
Data = Data & "," & iProperties.Value("Custom", "UsePartRev")

Data = Data & "," & SNFormat
Data = Data & "," & SNBaseDataType

'UOMClassID
Data = Data & "," & "BBN"

Data = Data & "," & SNMask
Data = Data & "," & SNMaskExample

'Net Weight UOM: only needed for manufactured parts
If StrComp(PartType, "M") = 0 Then
    Data = Data & "," & "LB"
Else
    Data = Data & ","
End If

'Write field headers & data to file
csv.WriteLine(Fields)
csv.WriteLine(Data)
csv.Close()

'Call the DMT on the generated CSV file
Dim dmt_loc = "C:\Epicor\ERP10.1Client\Client\DMT.exe"
Dim psi As New System.Diagnostics.ProcessStartInfo(dmt_loc)
psi.RedirectStandardOutput = True
psi.WindowStyle = ProcessWindowStyle.Hidden
psi.UseShellExecute = False

Dim username, password As String
username = "DMT_USERNAME"
password = "DMT_PASSWORD"
Dim connection As String = "net.tcp://CHERRY/EpicorPilot10"
psi.Arguments = "-NoUI=True -Import=""Part"" -Source=""" & FileName
psi.Arguments = psi.Arguments & """ -Add=True -Update=True -user=" & username
psi.Arguments = psi.Arguments & " -pass=" & password & " -ConnectionUrl="""
psi.Arguments = psi.Arguments & connection & """"

Dim dmt As System.Diagnostics.Process
dmt = System.Diagnostics.Process.Start(psi)
dmt.WaitForExit()

Dim msgSuccess As String = "Part successfully imported into Epicor!"
Dim msgFailure As String = "Error importing part into Epicor!"

Dim resultmsg As String
If dmt.ExitCode = 0 Then
    resultmsg = msgSuccess
Else
    resultmsg = msgFailure
End If

MsgBox(resultmsg)
