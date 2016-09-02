Imports Inventor

Public Module PartExport
    Sub Main()
    End Sub

    'create a CSV with the user-entered part data and run the DMT on it
    'returns:
    '   - 0 on success
    '   - 1 on fixable error
    '   - 2 on I/O error with log file
    '   - 3 on other error (see message box)
    '   - -1 on DMT timeout
    Public Function part_export(ByRef app As Inventor.Application, _
                                ByRef inv_params As UserParameters, _
                                ByRef dmt_obj As DMT) _
                                As Integer
        Dim fields, data As String
        Dim PartNum, SearchWord, Description, PartType, UOM As String
        Dim MfgComment, PurComment As String

        Dim inv_doc As Document = app.ActiveDocument
        Dim custom_props As PropertySet = inv_doc.PropertySets.Item("Inventor User Defined Properties")

        Dim part_entry As String = inv_params.Item("PartNumberToUse").Value
        Dim part_unpacked As Tuple(Of String, String, String) = SpeciesOps.unpack_pn(part_entry)

        'properties that will be used elsewhere, or need to be formatted for CSV
        PartNum = part_unpacked.Item1.ToUpper()
        Description = inv_params.Item("Description").Value
        SearchWord = Left(Description, 8)
        PartType = part_unpacked.Item2
        
        MfgComment = custom_props.Item("MfgComment").Value
        PurComment = custom_props.Item("PurComment").Value

        'UOM is set based on part type
        'NOTE: default is "M" (manufactured), though only "P"/"M" are expected
        If String.Equals(PartType, "P") Then
            UOM = "EAP"
        Else
            UOM = "EAM"
        End If

        Dim TrackSerialNum As Boolean = custom_props.Item("TrackSerialNum").Value

        'note: serial number fields may get appended
        fields = "Company,PartNum,SearchWord,PartDescription,ClassID,IUM,PUM,TypeCode,PricePerCode,ProdCode,MfgComment,PurComment,TrackSerialNum,SalesUM,UsePartRev,UOMClassID"

        'Build string containing values in order expected by DMT (see fields string)
        data = "BBN"                                'Company name (constant)
        data = data & "," & PartNum
        
        'Search word, first 8 characters of description
        data = data & "," & InventorOps.format_csv_field(SearchWord)
        data = data & "," & InventorOps.format_csv_field(Description)

        data = data & "," & custom_props.Item("ClassID").Value

        data = data & "," & UOM                     'IUM
        data = data & "," & UOM                     'PUM

        data = data & "," & PartType

        'Price per grouping (currently: "E", but will this always be the case?)
        data = data & "," & "E"

        data = data & "," & custom_props.Item("ProdCode").Value
        data = data & "," & InventorOps.format_csv_field(MfgComment)
        data = data & "," & InventorOps.format_csv_field(PurComment)
        data = data & "," & TrackSerialNum
        data = data & "," & UOM                     'SalesUM
        data = data & "," & custom_props.Item("UsePartRev").Value
        
        data = data & "," & "BBN"                   'UOMClassID

        'Net Weight UOM: only needed for manufactured parts
        If String.Equals(PartType, "M") Then
            fields = fields & ",NetWeightUOM"
            data = data & "," & "LB"
        End If

        'if serial number is being tracked, a bunch of fields are enabled
        If TrackSerialNum AndAlso String.Equals(PartType, "M") Then
            fields = fields & ",SNFormat,SNBaseDataType,SNMask,SNMaskExample"
            data = data & "," & "NF#######"
            data = data & "," & "MASK"
            data = data & "," & "NF"
            data = data & "," & "NF9999999"
        End If

        Dim file_name As String
        file_name = dmt_obj.write_csv("Part_Level.csv", fields, data)

        Return dmt_obj.dmt_import("Part", file_name, False)
    End Function
End Module
