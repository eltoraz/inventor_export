AddVbFile "inventor_common.vb"          'InventorOps.get_param_set

Imports Inventor

Public Class PartRevExport
    Sub Main()
    End Sub

    Public Shared Function part_rev_export(ByRef app As Inventor.Application, _
                                           ByRef dmt_obj As DMT)
        Dim fields, data As String
        Dim PartNum, RevisionNum, DrawNum As String
        Dim ApprovedDate As Date

        Dim inv_doc As Document = app.ActiveDocument
        Dim inv_params As UserParameters = InventorOps.get_param_set(app)
        Dim summary_props, design_props, custom_props As PropertySet

        summary_props = inv_doc.PropertySets.Item("Inventor Summary Information")
        design_props = inv_doc.PropertySets.Item("Design Tracking Properties")
        custom_props = inv_doc.PropertySets.Item("Inventor User Defined Properties")

        PartNum = inv_params.Item("PartNumberToUse").Value
        DrawNum = design_props.Item("Part Number").Value
        RevisionNum = summary_props.Item("Revision Number").Value
        ApprovedDate = design_props.Item("Engr Date Approved").Value

        fields = "Company,PartNum,RevisionNum,RevShortDesc,RevDescription,Approved,ApprovedDate,ApprovedBy,EffectiveDate,DrawNum,Plant,ProcessMode"

        data = "BBN"                        'Company name (constant)
        data = data & "," & PartNum
        data = data & "," & RevisionNum
        data = data & "," & "Revision " & RevisionNum
        data = data & "," & custom_props.Item("RevDescription").Value

        'Logic TODO: Approved hardcoded for now
        'Logic TODO: is there any reason for the user to specify EffectiveDate as
        '            anything different from ApprovedDate?
        data = data & "," & "True"          'Approved
        data = data & "," & ApprovedDate    'ApprovedDate
        data = data & "," & "d.laforce"     'ApprovedBy
        data = data & "," & ApprovedDate    'EffectiveDate

        data = data & "," & DrawNum         'DrawNum
        data = data & "," & "MfgSys"        'Plant (only one)
        data = data & "," & "S"             'ProcessMode (always sequential)

        Dim file_name As String
        file_name = dmt_obj.write_csv("Part_Rev.csv", fields, data)

        Return dmt_obj.dmt_import("Part Revision", file_name)
    End Function
End Class
