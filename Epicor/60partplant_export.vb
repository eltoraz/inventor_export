Imports Inventor

Public Class PartPlantExport
    Sub Main()
    End Sub

    Public Shared Function part_plant_export(ByRef app As Inventor.Application, _
                                             ByRef inv_params As userParameters, _
                                             ByRef dmt_obj As DMT)
        Dim fields, data As String
        Dim PartNum, PartType As String

        Dim inv_doc As Document = app.ActiveDocument
        Dim design_props, custom_props As PropertySet

        design_props = inv_doc.PropertySets.Item("Design Tracking Properties")
        custom_props = inv_doc.PropertySets.Item("Inventor User Defined Properties")

        Dim part_entry As String = inv_params.Item("PartNumberToUse").Value
        Dim part_unpacked As Tuple(Of String, String, String) = SpeciesOps.unpack_pn(part_entry)
        Dim pn As String = part_unpacked.Item1

        PartNum = pn

        'fields for manufactured parts
        Dim TrackSerialNumber As Boolean
        Dim SNMask, SNMaskExample, SNBaseDataType, SNFormat As String

        PartType = custom_props.Item("PartType").Value
        TrackSerialNum = custom_props.Item("TrackSerialNum").Value

        If TrackSerialNum AndAlso String.Equals(PartType, "M") Then
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
        data = data & "," & PartNum
        data = data & "," & "453"                       'PrimWhse (just one warehouse)

        'these fields won't get filled from Inventor
        data = data & "," & ""                          'LeadTime
        data = data & "," & ""                          'VendorNum
        data = data & "," & ""                          'PurPoint

        data = data & "," & PartType

        data = data & "," & "F"                         'CostMethod (constant)

        data = data & "," & SNMask
        data = data & "," & SNMaskExample
        data = data & "," & SNBaseDataType
        data = data & "," & SNFormat

        Dim file_name As String
        file_name = dmt_obj.write_csv("Part_Plant.csv", fields, data)

        Return dmt_obj.dmt_import("Part Plant", file_name, False)
    End Function
End Class
