Public Class Part_Rev_Export
    Public Shared Function Part_Rev()
        Dim fields, data As String
        Dim PartNum, RevisionNum As String
        Dim ApprovedDate As Date

        PartNum = iProperties.Value("Project", "Part Number")
        RevisionNum = iProperties.Value("Project", "Revision Number")
        ApprovedDate = iProperties.Value("Status", "Eng. Approved Date")

        fields = "Company,PartNum,RevisionNum,RevShortDesc,RevDescription,Approved,ApprovedDate,ApprovedBy,EffectiveDate,DrawNum,Plant,ProcessMode"

        data = "BBN"                        'Company name (constant)
        data = data & "," & PartNum
        data = data & "," & RevisionNum
        data = data & "," & "Revision " & RevisionNum
        data = data & "," & iProperties.Value("Custom", "RevDescription")

        'Logic TODO: Approved & ApprovedBy hardcoded for now
        'Logic TODO: is there any reason for the user to specify EffectiveDate as
        '            anything different from ApprovedDate?
        data = data & "," & "True"          'Approved
        data = data & "," & ApprovedDate    'ApprovedDate
        data = data & "," & "d.laforce"     'ApprovedBy
        data = data & "," & ApprovedDate    'EffectiveDate

        data = data & "," & PartNum         'DrawNum (same as part number)
        data = data & "," & "MfgSys"        'Plant (only one)
        data = data & "," & "S"             'ProcessMode (always sequential)

        Dim file_name As String
        file_name = DMT.write_csv("Part_Rev.csv", fields, data)

        Dim resultmsg As String = DMT.exec_DMT("Part Revision", file_name)
        Return resultmsg
    End Function
End Class
