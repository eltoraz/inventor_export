AddVbFile "dmt.vb"                      'DMT
AddVbFile "40part_export.vb"            'PartExport.part_export
AddVbFile "50partrev_export.vb"         'PartRevExport.part_rev_export
AddVbFile "60partplant_export.vb"       'PartPlantExport.part_plant_export
AddVbFile "species_list.vb"             'Species.species_list
AddVbFile "inventor_common.vb"          'InventorOps.get_param_set

Sub Main()
    'Pull latest data from Epicor
    'this data shouldn't change often, so the rule shouldn't need to be called often
    'DMT.dmt_export()

    'populate the PartNumberToUse param multi-value with the activated part numbers
    Dim app As Inventor.Application = ThisApplication
    Dim inv_params As UserParameters = InventorOps.get_param_set(app)

    'select the part we'll be working with here
    Dim active_parts As New ArrayList()
    Dim no_species As Boolean = False

    Try
        For Each s As String in Species.species_list
            Dim subst As String = Replace(s, "-", "4")

            Dim flag_param As Parameter = inv_params.Item("Flag" & subst)
            Dim flag_value = flag_param.Value

            If flag_value Then
                'add active parts and materials to the list to present to the user
                Dim part_param As Parameter = inv_params.Item("Part" & subst)
                Dim part_value As String = part_param.Value
                active_parts.Add(part_value)

                If StrComp(s, "Hardware") <> 0 Then
                    Dim mat_param As Parameter = inv_params.Item("Mat" & subst)
                    Dim mat_value As String = mat_param.Value
                    active_parts.Add(mat_value)
                End If
            End If
        Next
    Catch e As ArgumentException
        'exception thrown when using UserParameters.Item above and trying to
        'get a parameter that doesn't exist
        no_species = True
    End Try

    'can't proceed if there isn't a part number for at least one species
    If no_species OrElse active_parts.Count = 0 Then
        MsgBox("Warning: there are no species defined for this part. Please " & _
               "run the BBN Species Setup first.")
        Return
    End If
    MultiValue.List("PartNumberToUse") = active_parts

    'TODO: allow cancelling form here to abort
    Dim part_selected As Boolean = False
    Dim pn As String = ""
    Do
        iLogicForm.ShowGlobal("epicor_15part_select", FormMode.Modal)

        pn = inv_params.Item("PartNumberToUse").Value
        If StrComp(pn, "") <> 0 Then
            part_selected = True
        Else
            MsgBox("Please select a part to continue with the Epicor export.")
            iLogicVb.RunExternalRule("dummy.vb")
        End If
    Loop While Not part_selected

    'Call the other rules in order
    iLogicVb.RunExternalRule("10multi_value.vb")
    iLogicForm.ShowGlobal("epicor_20part_properties", FormMode.Modal)
    iLogicVb.RunExternalRule("25logic_check.vb")
    iLogicVb.RunExternalRule("30set_props.vb")

    'if part export fails, abort - this will usually mean the part is already
    'in the DB and so the straight add operation failed
    Dim dmt_obj As New DMT()
    Dim ret_value = PartExport.part_export(ThisApplication, dmt_obj)
    If ret_value = 0 Then
        PartRevExport.part_rev_export(ThisApplication, dmt_obj)
        PartPlantExport.part_plant_export(ThisApplication, dmt_obj)
    ElseIf ret_value = -1 Then
        MsgBox("Error: DMT timed out. Aborting...")
    Else
        MsgBox("Warning: this part is already present in Epicor. Aborting...")
    End If

    'TODO: display message box about DMT state - maybe last 3 lines of logfile
End Sub
