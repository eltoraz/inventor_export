AddVbFile "dmt.vb"

Public Class Part_Export
    Public Shared Function Part()
        Dim fields, data As String
        Dim Description, PartType, UOM As String
        Dim TrackSerialNum As Boolean
        Dim SNFormat, SNBaseDataType, SNMask, SNMaskExample As String

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
        file_name = DMT.write_csv("Part_Level.csv", fields, data)

        Dim resultmsg As String = DMT.exec_DMT("Part", file_name)
        Return resultmsg
    End Function
End Class
