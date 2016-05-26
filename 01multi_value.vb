AddVbFile "dmt.vb"

'create parameters with a restricted set of accepted values for import by
'Epicor DMT (the actual IDs the tool needs are set in set_props.vb)
Sub Main()
    Dim params As New Dictionary(Of String, UnitsTypeEnum)
    'Part parameters
    params.Add("PartType", UnitsTypeEnum.kTextUnits)
    params.Add("ProdCode", UnitsTypeEnum.kTextUnits)
    params.Add("ClassID", UnitsTypeEnum.kTextUnits)
    params.Add("UsePartRev", UnitsTypeEnum.kBooleanUnits)
    params.Add("MfgComment", UnitsTypeEnum.kTextUnits)
    params.Add("PurComment", UnitsTypeEnum.kTextUnits)
    params.Add("TrackSerialNum", UnitsTypeEnum.kBooleanUnits)

    'Revision parameters
    params.Add("RevDescription", UnitsTypeEnum.kTextUnits)
    
    'Plant parameters
    params.Add("LeadTime", UnitsTypeEnum.kUnitlessUnits)

    For Each kvp As KeyValuePair(Of String, UnitsTypeEnum) in params
        createParam(kvp.Key, kvp.Value)
    Next

    MultiValue.SetList("PartType", "M", "P")
    MultiValue.SetList("ProdCode", "FSC Assemblies", "FSC Components", "NCA Assemblies", "NCA Components", "Purchases")
    MultiValue.SetList("ClassID", "Assemblies we sell", "Box Materials", "Components we sell", "Finished Components for kits", "Finish Materials", "FSC Lumber", "Finished Components on shelf", "IT Supplies", "Lumber", "Office Supplies", "Other Materials", "Tooling")
    'TODO: multi-value for approving engineer for revision?

    'TODO: multi-value for vendors for purchased parts?
End Sub

Sub createParam(ByVal n As String, ByVal paramType As UnitsTypeEnum)
    dim invDoc As Document = ThisApplication.ActiveDocument

    Dim invParams As UserParameters = invDoc.Parameters.UserParameters

    Dim TestParam As UserParameter

    'if the parameter doesn't already exist, UserParameters.Item will throw an error
    Try
        TestParam = invParams.Item(n)
    Catch
        Dim defaultValue
        If paramType = UnitsTypeEnum.kTextUnits Then
            defaultValue = ""
        ElseIf paramType = UnitsTypeEnum.kBooleanUnits Then
            defaultValue = True
        ElseIf paramType = UnitsTypeEnum.kUnitlessUnits Then
            defaultValue = 0
        End If

        TestParam = invParams.AddByValue(n, defaultValue, paramType)
        invDoc.Update
    End Try
End Sub

'populate the list of options for parameter `n` from the CSV file `f`
'CSV should be populated by dmt.vb's dmt_export() method
Sub populate_multivalue(ByVal n As String, ByVal f As String)
    Dim file_name As String = DMT.dmt_working_path & "ref\" & f

    Using csv_reader As New Microsoft.VisualBasic.FileIO.TextFieldParser(file_name)
        csv_reader.TextFieldType = FileIO.FieldType.Delimited
        csv_reader.SetDelimiters(",")

        Dim current_row As String()
        While Not csv_reader.EndOfData
            'TODO
            Try
                'TODO
                current_row = csv_reader.ReadFields()
            End Try
        End While
    End Using
End Sub
