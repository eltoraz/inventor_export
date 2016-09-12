AddVbFile "inventor_common.vb"      'InventorOps.update_prop
AddVbFile "parameters.vb"           'ParameterOps.get_param_set, species_list

'create/update iProperties with the values entered in form 30 (enabled in form 20)
Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim inv_doc As Document = app.ActiveEditDocument
    Dim inv_params As UserParameters = ParameterOps.get_param_set(app)

    Dim materials_only As Boolean = inv_params.Item("MaterialsOnly").Value
    Dim is_part_doc As Boolean = TypeOf inv_doc Is PartDocument

    'loop through each species in the list, but process
    ' only those whose flag is enabled
    For Each s As String In ParameterOps.species_list
        Dim subst As String = Replace(s, "-", "4")

        'note: "Hardware" doesn't have a part and thus plain Flag parameter associated
        ' but every species has a FlagMat parameter
        Dim flag_value As Boolean
        If String.Equals(s, "Hardware") Then
            flag_value = False
        Else
            flag_value = inv_params.Item("Flag" & subst).Value
        End If
        Dim mat_flag_value As Boolean = inv_params.Item("FlagMat" & subst).Value

        'part (convert lower-case to upper on the way too)
        'note: "Hardware" is a material category and doesn't have a Part associated
        ' but don't need to check for it here since the flag's set to False above in that case
        If flag_value AndAlso Not materials_only Then
            Dim part_param As Parameter = inv_params.Item("Part" & subst)
            Dim part_value As String = part_param.Value.ToUpper()
            InventorOps.update_prop("Part (" & s & ")", part_value, app)
        End If

        'material: skip for assemblies
        If mat_flag_value AndAlso is_part_doc Then
            Dim mat_param As Parameter = inv_params.Item("Mat" & subst)
            Dim mat_value As String = mat_param.Value.ToUpper()
            InventorOps.update_prop("Material (" & s & ")", mat_value, app)
        End If
    Next
End Sub
