AddVbFile "dmt.vb"                  'DMT.dmt_working_path
AddVbFile "parameters.vb"           'ParameterOps.create_all_params

'create parameters with a restricted set of accepted values for import by
'Epicor DMT (the actual IDs the tool needs are set in set_props.vb)
Sub Main()
    Dim inv_app As Inventor.Application = ThisApplication

    ParameterOps.create_all_params(inv_app)

    MultiValue.SetList("PartType", "M", "P")
    MultiValue.List("ProdCode") = fetch_list_values("ProdCode.csv")
    MultiValue.List("ClassID") = fetch_list_values("ClassID.csv")

    'TODO: multi-value for approving engineer for revision?
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
