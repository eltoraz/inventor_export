Imports Inventor

Public Module PartPlantExport
    Sub Main()
    End Sub

    Public Function part_plant_export(ByRef app As Inventor.Application, _
                                      ByRef inv_params As userParameters, _
                                      ByRef dmt_obj As DMT) _
                                      As Integer
        Dim fields, data As String
        Dim PartNum, PartType As String

        Dim inv_doc As Document = app.ActiveDocument
        Dim design_props, custom_props As PropertySet

        design_props = inv_doc.PropertySets.Item("Design Tracking Properties")
        custom_props = inv_doc.PropertySets.Item("Inventor User Defined Properties")

        Dim part_entry As String = inv_params.Item("PartNumberToUse").Value
        Dim part_unpacked As Tuple(Of String, String, String) = SpeciesOps.unpack_pn(part_entry)

        PartNum = part_unpacked.Item1.ToUpper()
        PartType = part_unpacked.Item2

        'fields for manufactured parts
        Dim SNMask, SNMaskExample, SNBaseDataType, SNFormat As String

        fields = "Company,Plant,PartNum,PrimWhse,SourceType,CostMethod"

        data = "BBN"                                    'Company name (constant)
        data = data & "," & "MfgSys"                    'Plant (only one for this company)
        data = data & "," & PartNum
        data = data & "," & "453"                       'PrimWhse (just one warehouse)

        data = data & "," & PartType

        data = data & "," & "F"                         'CostMethod (constant)

        Dim TrackSerialNum As Boolean = custom_props.Item("TrackSerialNum").Value
        If TrackSerialNum AndAlso String.Equals(PartType, "M") Then
            fields = fields & ",SNMask,SNMaskExample,SNBaseDataType,SNFormat"
            data = data & "," & "NF"
            data = data & "," & "NF9999999"
            data = data & "," & "MASK"
            data = data & "," & "NF#######"
        End If

        Dim file_name As String
        file_name = dmt_obj.write_csv("Part_Plant.csv", fields, data)

        Return dmt_obj.dmt_import("Part Plant", file_name, True)
    End Function
End Module
