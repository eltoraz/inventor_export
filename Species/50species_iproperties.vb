AddVbFile "inventor_common.vb"      'InventorOps.update_prop, get_param_set
AddVbFile "species_list.vb"         'Species.species_list

'create/update iProperties with the values entered in form 30 (enabled in form 20)
Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim inv_params As UserParameters = InventorOps.get_param_set(app)

    For Each s As String In Species.species_list
        Dim subst As String = Replace(s, "-", "4")

        Dim flag_param As Parameter = inv_params.Item("Flag" & subst)
        Dim flag_value = flag_param.Value

        If flag_value Then
            'part (convert lower-case to upper on the way too)
            Dim part_param As Parameter = inv_params.Item("Part" & subst)
            Dim part_value As String = part_param.Value.ToUpper()
            InventorOps.update_prop("Part (" & s & ")", part_value, inv_app)

            'material: skip for "Hardware"
            If StrComp(s, "Hardware") <> 0 Then
                Dim mat_param As Parameter = inv_params.Item("Mat" & subst)
                Dim mat_value As String = mat_param.Value.ToUpper()
                InventorOps.update_prop("Material (" & s & ")", mat_value, inv_app)
            End If
        End If
    Next
End Sub
