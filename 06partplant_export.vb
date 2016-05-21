AddVbFile "dmt.vb"

Sub Main()
    Dim fields, data As String
    Dim PartType As String

    'fields for purchased parts
    Dim LeadTime, VendorNum As Integer
    Dim PurPoint As String

    'fields for manufactured parts
    Dim TrackSerialNumber As Boolean
    Dim SNMask, SNMaskExample, SNBaseDataType, SNFormat As String

    PartType = iProperties.Value("Custom", "PartType")
    TrackSerialNum = iProperties.Value("Custom", "TrackSerialNum")

    If StrComp(PartType, "P") = 0 Then
        'TODO: setup parameters/iProperties for LeadTime et al.
        'TODO: come to think of it, will purchased parts go through Inventor at all?
    Else
        LeadTime = 0
        VendorNum = 0
        PurPoint = ""
    End If

    If TrackSerialNumber And StrComp(PartType, "M") = 0 Then
        SNMask = "NF"
        SNMaskExample = "NF9999999"
        SNBaseDataType = "MASK"
        SNFormat = "NF#######"
    Else
        SNMask = ""
        SNMaskExample = ""
        SNBaseDataType = ""
        SNFormat = ""
    End If

    fields = "Company,Plant,PartNum,PrimWhse,LeadTime,VendorNum,PurPoint,SourceType,CostMethod,SNMask,SNMaskExample,SNBaseDataType,SNFormat"

    data = "BBN"                                    'Company name (constant)
    data = data & "," & "MfgSys"                    'Plant (only one for this company)
    data = data & "," & iProperties.Value("Project", "Part Number')
    data = data & "," & "453"                       'PrimWhse (just one warehouse)

    data = data & "," & LeadTime
    data = data & "," & VendorNum
    data = data & "," & PurPoint

    data = data & "," & PartType

    data = data & "," & "F"                         'CostMethod (constant)

    data = data & "," & SNMask
    data = data & "," & SNMaskExample
    data = data & "," & SNBaseDataType
    data = data & "," & SNFormat

    Dim file_name As String
    file_name = DMT.write_csv("Part_Plant.csv", fields, data)

    'TODO: verify this is the correct table name in DMT
    DMT.exec_DMT("Part Plant", file_name)
End Sub