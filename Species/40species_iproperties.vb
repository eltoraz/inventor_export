AddVbFile "inventor_common.vb"      'InventorOps.update_prop

'create/update iProperties with the values entered in form 30 (enabled in form 20)
Sub Main()
    Dim inv_app As Inventor.Application = ThisApplication
    Dim inv_doc As Document = inv_app.ActiveDocument
    Dim inv_params As UserParameters = inv_doc.Parameters.UserParameters

    For i = 1 To inv_params.Count
        Dim param As Parameter = inv_params.Item(i)
        Dim param_name As String = param.Name
        Dim param_value = param.Value

        'TODO: validate that the corresponding "Part"/"Mat" parameters have
        '      been populated if the flag is true, and also that they have the
        '      correct formatting (XX-###)
        If StrComp(Left(param_name, 4), "Flag") = 0 AndAlso param_value = True Then
            'get the proper species name (remove "Flag" and replace placeholder "4")
            Dim param_species As String = Left(param_name, 4)
            Dim species_name As String = Replace(param_name, "4", "-").Substring(4)

            'part
            Dim part_param As Parameter = inv_params.Item("Part" & param_species)
            Dim part_value As String = part_param.Value
            InventorOps.update_prop("Part (" & species_name & ")", part_value, inv_app)

            'material: skip for "Hardware"
            If StrComp(species_name, "Hardware") <> 0 Then
                Dim mat_param As Parameter = inv_params.Item("Mat" & param_species)
                Dim mat_value As String = mat_param.Value
                InventorOps.update_prop("Material (" & species_name & ")", mat_value, inv_app)
            End If
        End If
    Next
End Sub
