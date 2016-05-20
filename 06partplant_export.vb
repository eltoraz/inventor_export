AddVbFile "dmt.vb"

Sub Main()
    Dim fields, data As String
    Dim PartType As String
    Dim LeadTime, VendorNum, PurPoint As String     'only valid for purchased parts

    PartType = iProperties.Value("Custom", "PartType")

    If StrComp(PartType, "P") = 0 Then
        'TODO: setup parameters/iProperties for LeadTime et al.
    Else
        'TODO: setup serial number for manufactured parts
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

    'TODO: Serial Number
End Sub
