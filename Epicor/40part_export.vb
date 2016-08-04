Imports Inventor

Public Class PartExport
    Sub Main()
    End Sub

    Public Shared Function part_export(ByRef app As Inventor.Application, _
                                       ByRef inv_params As UserParameters, _
                                       ByRef dmt_obj As DMT)
        Dim fields, data As String
        Dim PartNum, SearchWord, Description, PartType, UOM As String
        Dim MfgComment, PurComment As String
        Dim TrackSerialNum As Boolean
        Dim SNFormat, SNBaseDataType, SNMask, SNMaskExample As String

        Dim inv_doc As Document = app.ActiveDocument
        Dim custom_props As PropertySet = inv_doc.PropertySets.Item("Inventor User Defined Properties")

        Dim part_entry As String = inv_params.Item("PartNumberToUse").Value
        Dim part_unpacked As Tuple(Of String, String, String) = SpeciesOps.unpack_pn(part_entry)
        Dim pn As String = part_unpacked.Item1

        'properties that will be used elsewhere, or need to be formatted for CSV
        PartNum = pn
        Description = inv_params.Item("Description").Value
        SearchWord = Left(Description, 8)
        PartType = part_unpacked.Item2
        
        MfgComment = custom_props.Item("MfgComment").Value
        PurComment = custom_props.Item("PurComment").Value

        'UOM is set based on part type
        'NOTE: default is "M" (manufactured), though only "P"/"M" are expected
        If StrComp(PartType, "P") = 0 Then
            UOM = "EAP"
        Else
            UOM = "EAM"
        End If

        TrackSerialNum = custom_props.Item("TrackSerialNum").Value

        'if serial number is being tracked, a bunch of fields are enabled
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
        data = data & "," & PartNum
        
        'Search word, first 8 characters of description
        data = data & "," & InventorOps.format_csv_field(SearchWord)
        data = data & "," & InventorOps.format_csv_field(Description)

        data = data & "," & custom_props.Item("ClassID").Value

        data = data & "," & UOM
        data = data & "," & UOM

        data = data & "," & PartType

        'Price per grouping (currently: "E", but will this always be the case?)
        data = data & "," & "E"

        data = data & "," & custom_props.Item("ProdCode").Value
        data = data & "," & InventorOps.format_csv_field(MfgComment)
        data = data & "," & InventorOps.format_csv_field(PurComment)
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

        Return dmt_obj.dmt_import("Part", file_name, False)
    End Function
End Class
