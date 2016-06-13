Imports Inventor

Public Class PartExport
    Sub Main()
    End Sub

    Public Shared Function part_export(ByRef app As Inventor.Application, ByRef dmt_obj As DMT)
        Dim fields, data As String
        Dim Description, PartType, UOM As String
        Dim TrackSerialNum As Boolean
        Dim SNFormat, SNBaseDataType, SNMask, SNMaskExample As String

        Dim inv_doc As Document = app.ActiveDocument
        Dim design_props, custom_props As PropertySet

        design_props = inv_doc.PropertySets.Item("Design Tracking Properties")
        custom_props = inv_doc.PropertySets.Item("Inventor User Defined Properties")

        Description = design_props.Item("Description").Value
        PartType = custom_props.Item("PartType").Value

        'UOM is set based on part type
        'NOTE: default is "M" (manufactured), though only "P"/"M" are expected
        If StrComp(PartType, "P") = 0 Then
            UOM = "EAP"
        Else
            UOM = "EAM"
        End If

        TrackSerialNum = custom_props.Item("TrackSerialNum").Value

        'if serial number is being tracked, a bunch of fields are enabled
        'logic TODO: verify this works
        If TrackSerialNum AndAlso StrComp(PartType, "M") = 0 Then
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
        'TODO: this is actuall the drawing number, and the part number will be
        '      found in a custom iProperty populated by the species chooser
        data = data & "," & design_props.Item("Part Number").Value

        data = data & "," & Left(Description, 8)    'Search word, first 8 characters of description
        data = data & "," & Description

        data = data & "," & custom_props.Item("ClassID").Value

        data = data & "," & UOM
        data = data & "," & UOM

        data = data & "," & PartType

        'Price per grouping (currently: "E", but will this always be the case?)
        data = data & "," & "E"

        data = data & "," & custom_props.Item("ProdCode").Value
        data = data & "," & custom_props.Item("MfgComment").Value
        data = data & "," & custom_props.Item("PurComment").Value
        data = data & "," & TrackSerialNum
        data = data & "," & UOM
        data = data & "," & custom_props.Item("UsePartRev").Value

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
        file_name = dmt_obj.write_csv("Part_Level.csv", fields, data)

        Return dmt_obj.dmt_import("Part", file_name)
    End Function
End Class
