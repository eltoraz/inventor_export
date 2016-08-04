AddVbFile "dmt.vb"                  'DMT.dmt_working_path
AddVbFile "inventor_common.vb"      'InventorOps.update_prop, get_param_set
AddVbFile "parameters.vb"           'ParameterLists.epicor_params

'set iProperties with values the user has defined in a form
'note: these values will mostly be the IDs the Epicor DMT is expecting rather
'      than the human-readable strings
Sub Main()
    'list of parameters that need to be converted to custom iProperties
    Dim app As Inventor.Application = ThisApplication
    Dim inv_doc As Document = app.ActiveDocument
    Dim inv_params As UserParameters = InventorOps.get_param_set(app)

    'mappings for human-readable values (i.e. in the dropdown boxes) -> keys
    'only necessary for ProdCode and ClassID
    Dim ProdCodeMap As Dictionary(Of String, String) = fetch_list_mappings("ProdCode.csv")
    Dim ClassIDMap As Dictionary(Of String, String) = fetch_list_mappings("ClassID.csv")

    'TODO: map approving engineers to Epicor IDs?

    For Each kvp As KeyValuePair(Of String, UnitsTypeEnum) in ParameterLists.epicor_params
        'if Epicor requires a short ID, convert the human-readable value via
        'the appropriate mapping (see above)
        'required for: ProdCode, ClasID
        Dim param As Parameter = inv_params.Item(kvp.Key)
        Dim param_name As String = param.Name
        Dim param_value = param.Value

        If String.Equals(param_name, "ProdCode") Then
            param_value = ProdCodeMap(param_value)
        Else If String.Equals(param_name, "ClassID") Then
            param_value = ClassIDMap(param_value)
        Else If String.Equals(param_name, "MfgComment") Then
            'note: Epicor comment fields support up to 16000 chars
            param_value = Left(param_value, 16000)
        Else If String.Equals(param_name, "PurComment") Then
            param_value = Left(param_value, 16000)
        Else If String.Equals(param_name, "RevDescription") Then
            param_value = Left(param_value, 16000)
        End If

        InventorOps.update_prop(param_name, param_value, app)

        inv_doc.Update
    Next
End Sub

'Map the description in the parameter to the DB friendly ID expected by Epicor
Function fetch_list_mappings(ByVal f As String) As Dictionary(Of String, String)
    Dim file_name As String = DMT.dmt_working_path & "ref\" & f
    Dim mapping As New Dictionary(Of String, String)

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
                'skip headers
                first_line = False
            Else
                mapping.Add(current_row(0), current_row(1))
            End If
        End While
    End Using

    Return mapping
End Function
