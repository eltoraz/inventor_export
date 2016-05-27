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
    params.Add("VendorNum", UnitsTypeEnum.kTextUnits)
    'vendor parameter will pass a number, but the selection box needs to be
    'be human-readable and not just an ID
    'TODO: purchase point

    For Each kvp As KeyValuePair(Of String, UnitsTypeEnum) in params
        createParam(kvp.Key, kvp.Value)
    Next

    MultiValue.SetList("PartType", "M", "P")
    MultiValue.List("ProdCode") = fetch_list_values("ProdCode.csv")
    MultiValue.List("ClassID") = fetch_list_values("ClassID.csv")
    
    MultiValue.List("VendorNum") = fetch_list_values("VendorNum.csv")

    'TODO: multi-value for approving engineer for revision?
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

'populate the list of options from the CSV file specified by `f`
'(found in `DMT.dmt_working_path`\ref)
'the CSV should in turn be populated by dmt.vb's DMT.dmt_export() method
Function fetch_list_values(ByVal f As String) As ArrayList
    Dim file_name As String = DMT.dmt_working_path & "ref\" & f
    Dim option_list As New ArrayList()

    Using csv_reader As New FileIO.TextFieldParser(file_name)
        csv_reader.TextFieldType = FileIO.FieldType.Delimited
        csv_reader.SetDelimiters(",")

        Dim current_row As String()
        Dim first_line As Boolean = True
        While Not csv_reader.EndOfData
            Try
                current_row = csv_reader.ReadFields()
            Catch ex As FileIO.MalformedLineException
                Debug.Write("CSV contained invalid line:" & ex.Message)
            End Try

            If first_line Then
                'skip the header row
                first_line = False
            Else
                'we only need the first field here, which the queries export
                'as the human-readable descriptions
                option_list.Add(current_row(0))
            End If
        End While
    End Using

    Return option_list
End Function
