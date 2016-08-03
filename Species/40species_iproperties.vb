AddVbFile "inventor_common.vb"      'InventorOps.update_prop, get_param_set
AddVbFile "species_list.vb"         'Species.species_list

'create/update iProperties with the values entered in form 30 (enabled in form 20)
Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim inv_doc As Document = app.ActiveEditDocument
    Dim inv_params As UserParameters = InventorOps.get_param_set(app)

    Dim materials_only As Boolean = inv_params.Item("MaterialsOnly").Value
    Dim is_part_doc As Boolean = TypeOf inv_doc Is PartDocument

    For Each s As String In Species.species_list
        Dim subst As String = Replace(s, "-", "4")

        Dim flag_param As Parameter = inv_params.Item("Flag" & subst)
        Dim flag_value = flag_param.Value

        If flag_value Then
            'part (convert lower-case to upper on the way too)
            If Not materials_only Then
                Dim part_param As Parameter = inv_params.Item("Part" & subst)
                Dim part_value As String = part_param.Value.ToUpper()
                InventorOps.update_prop("Part (" & s & ")", part_value, app)
            End If

            'material: skip for "Hardware"
            If is_part_doc AndAlso StrComp(s, "Hardware") <> 0 Then
                Dim mat_param As Parameter = inv_params.Item("Mat" & subst)
                Dim mat_value As String = mat_param.Value.ToUpper()
                InventorOps.update_prop("Material (" & s & ")", mat_value, app)
            End If
        End If
    Next
End Sub
