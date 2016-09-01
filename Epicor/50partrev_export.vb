Imports System.Text.RegularExpressions
Imports Inventor

Public Module PartRevExport
    Sub Main()
    End Sub

    'create a CSV with the details for this part revision and export it via DMT
    'returns:
    '   - 0 on success
    '   - 1 on fixable error
    '   - 2 on I/O error with log file
    '   - 3 on other error (see message box)
    '   - -1 on DMT timeout
    Public Function part_rev_export(ByRef app As Inventor.Application, _
                                    ByRef inv_params As UserParameters, _
                                    ByRef dmt_obj As DMT) _
                                    As Integer
        Dim fields, data As String
        Dim PartNum, RevisionNum, RevDescription, DrawNum, UserName As String
        Dim ApprovedDate As Date

        Dim inv_doc As Document = app.ActiveDocument
        Dim summary_props, design_props, custom_props As PropertySet

        'iProperties this module uses are spread amongst several sets
        summary_props = inv_doc.PropertySets.Item("Inventor Summary Information")
        design_props = inv_doc.PropertySets.Item("Design Tracking Properties")
        custom_props = inv_doc.PropertySets.Item("Inventor User Defined Properties")

        Dim part_entry As String = inv_params.Item("PartNumberToUse").Value
        Dim part_unpacked As Tuple(Of String, String, String) = SpeciesOps.unpack_pn(part_entry)

        'somewhat confusingly, design team puts the drawing number in the part number property
        'a drawing will have multiple part numbers associated: one for each species it will
        ' be made in, and one for the associated raw materials
        PartNum = part_unpacked.Item1.ToUpper()
        DrawNum = design_props.Item("Part Number").Value
        RevisionNum = summary_props.Item("Revision Number").Value.ToUpper()
        RevDescription = custom_props.Item("RevDescription").Value
        ApprovedDate = design_props.Item("Engr Date Approved").Value

        'set username to use in approving engineer field
        'NOTE: this uses the Windows username unless it doesn't match the standard,
        '      so it's pretty fragile; there are just a couple cases in BBN where
        '      the user needs to manually enter their name (otherwise the username is
        '      the same in both Active Directory and Epicor)
        UserName = System.Environment.UserName
        Dim regex_match As Match = New Regex("^\w\.\w+$").Match(UserName)
        If Not regex_match.Success Then
            UserName = InputBox("Please enter your Epicor username:", "Epicor username")
        End If

        fields = "Company,PartNum,RevisionNum,RevShortDesc,RevDescription,Approved,ApprovedDate,ApprovedBy,EffectiveDate,DrawNum,Plant,ProcessMode"

        data = "BBN"                        'Company name (constant)
        data = data & "," & PartNum
        data = data & "," & RevisionNum
        data = data & "," & InventorOps.format_csv_field("Revision " & RevisionNum)
        data = data & "," & InventorOps.format_csv_field(RevDescription)

        'Logic TODO: Approved hardcoded for now
        'Logic TODO: is there any reason for the user to specify EffectiveDate as
        '            anything different from ApprovedDate?
        data = data & "," & "True"          'Approved
        data = data & "," & ApprovedDate    'ApprovedDate
        data = data & "," & UserName        'ApprovedBy
        data = data & "," & ApprovedDate    'EffectiveDate

        data = data & "," & DrawNum         'DrawNum
        data = data & "," & "MfgSys"        'Plant (only one)
        data = data & "," & "S"             'ProcessMode (always sequential)

        Dim file_name As String
        file_name = dmt_obj.write_csv("Part_Rev.csv", fields, data)

        Return dmt_obj.dmt_import("Part Revision", file_name, False)
    End Function
End Module
