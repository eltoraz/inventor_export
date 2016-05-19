AddVbFile "dmt.vb"

Sub Main()
    Dim dmt_log As String = ""

    dmt_log = dmt_log & Part(csv_path) & Environment.NewLine
    MsgBox(dmt_log)
End Sub

Function Part(csv_path As String)
    Dim fields, data
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

    fields = "Company,PartNum,SearchWord,PartDescription,ClassID,IUM,PUM,TypeCode,PricePerCode,ProdCode,MfgComment,PurComment,TrackSerialNum,SalesUM,UsePartRev,SNFormat,SNBaseDataType,UOMClassID,SNMask,SNMaskExample,NetWeightUOM"

    'Build string containing values in order expected by DMT (see fields string)
    data = "BBN"                                'Company name (constant)
    data = data & "," & iProperties.Value("Project", "Part Number")

    data = data & "," & Left(Description, 8)    'Search word, first 8 characters of description
    data = data & "," & Description

    data = data & "," & iProperties.Value("Custom", "ClassID")

    data = data & "," & UOM
    data = data & "," & UOM

    data = data & "," & PartType

    'Price per grouping (currently: "E", but will this always be the case?)
    data = data & "," & "E"

    data = data & "," & iProperties.Value("Custom", "ProdCode")
    data = data & "," & iProperties.Value("Custom", "MfgComment")
    data = data & "," & iProperties.Value("Custom", "PurComment")
    data = data & "," & TrackSerialNum
    data = data & "," & UOM
    data = data & "," & iProperties.Value("Custom", "UsePartRev")

    data = data & "," & SNFormat
    data = data & "," & SNBaseDataType

    'UOMClassID
    data = data & "," & "BBN"

    data = data & "," & SNMask
    data = data & "," & SNMaskExample

    'Net Weight UOM: only needed for manufactured parts
    If StrComp(PartType, "M") = 0 Then
        data = data & "," & "LB"
    Else
        data = data & ","
    End If

    Dim file_name As String
    file_name = write_csv("Part_Level.csv", fields, data)

    Dim resultmsg As String = DMT.exec_DMT("Part", file_name)
    Return resultmsg
End Function

Function Part_Rev(csv_path As String)
    Dim fields, data
    Dim PartNum, RevisionNum

    PartNum = iProperties.Value("Project", "Part Number")
    RevisionNum = iProperties.Value("Project", "Revision Number")

    fields = "Company,PartNum,RevisionNum,RevShortDesc,RevDescription,Approved,ApprovedDate,ApprovedBy,EffectiveDate,DrawNum,Plant,ProcessMode"

    data = "BBN"                        'Company name (constant)
    data = data & "," & PartNum
    data = data & "," & RevisionNum
    data = data & "," & "Revision " & RevisionNum
    data = data & "," & iProperties.Value("Custom", "RevDescription")

    data = data & "," & "True"          'Approved (always approved?)
    data = data & "," & ""              'TODO: ApprovedDate
    data = data & "," & ""              'TODO: ApprovedBy
    data = data & "," & ""              'TODO: EffectiveDate

    data = data & "," & PartNum         'DrawNum (same as part number)
    data = data & "," & "MfgSys"        'Plant (only one)
    data = data & "," & "S"             'ProcessMode (always sequential)

    Dim file_name As String
    file_name = write_csv("Part_Rev.csv", fields, data)

    Dim resultmsg As String = DMT.exec_DMT("Part Revision", file_name)
End Function

Function write_csv(csv_name As String, fields As String, data As String)
    Dim fso, file_name, csv
    Dim csv_path As String = "I:\Cadd\_iLogic\Export\"

    'Open the CSV file (note: this will overwrite the file if it exists!)
    fso = CreateObject("Scripting.FileSystemObject")
    file_name = csv_path & csv_name
    csv = fso.OpenTextFile(file_name, 2, True, -2)

    'Write field headers & data to file
    csv.WriteLine(fields)
    csv.WriteLine(data)
    csv.Close()

    'need to return the full path & filename to pass to DMT
    Return file_name
End Function
