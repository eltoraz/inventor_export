Imports Inventor

Public Class PartPlantExport
    Sub Main()
    End Sub

    Public Shared Function part_plant_export(ByRef app As Inventor.Application, _
                                             ByRef dmt_obj As DMT)
        Dim fields, data As String
        Dim PartType As String

        Dim inv_doc As Document = app.ActiveDocument
        Dim design_props, custom_props As PropertySet

        design_props = inv_doc.PropertySets.Item("Design Tracking Properties")
        custom_props = inv_doc.PropertySets.Item("Inventor User Defined Properties")

        'fields for purchased parts
        Dim LeadTime, VendorNum, PurPoint As String

        'fields for manufactured parts
        Dim TrackSerialNumber As Boolean
        Dim SNMask, SNMaskExample, SNBaseDataType, SNFormat As String

        PartType = custom_props.Item("PartType").Value
        TrackSerialNum = custom_props.Item("TrackSerialNum").Value

        'fields that won't get filled when making the parts in Inventor
        LeadTime = ""
        VendorNum = ""
        PurPoint = ""

        If TrackSerialNum AndAlso StrComp(PartType, "M") = 0 Then
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
        'TODO: this is the drawing number, need to get the actual part number
        '      once the species populator is finished
        data = data & "," & design_props.Item("Part Number").Value
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
        file_name = dmt_obj.write_csv("Part_Plant.csv", fields, data)

        Return dmt_obj.dmt_import("Part Plant", file_name)
    End Function
End Class
