AddVbFile "inventor_common.vb"      'InventorOps.update_prop
AddVbFile "species_list.vb"         'Species.species_list

'create/update iProperties with the values entered in form 30 (enabled in form 20)
Sub Main()
    Dim inv_app As Inventor.Application = ThisApplication
    Dim inv_doc As Document = inv_app.ActiveEditDocument
    Dim part_doc As PartDocument
    Dim assm_doc As AssemblyDocument
    Dim inv_params As UserParameters

    If TypeOf inv_doc Is PartDocument Then
        part_doc = inv_app.ActiveEditDocument
        inv_params = part_doc.ComponentDefinition.Parameters.UserParameters
    ElseIf TypeOf inv_doc Is AssemblyDocument Then
        assm_doc = inv_app.ActiveEditDocument
        inv_params = assm_doc.ComponentDefinition.Parameters.UserParameters
    Else
        inv_params = inv_doc.ComponentDefinition.Parameters.UserParameters
    End If

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
