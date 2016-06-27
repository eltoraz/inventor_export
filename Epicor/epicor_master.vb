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

    Dim form_result As FormResult

    Do
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
        
        'check whether the species parameters have been created but none selected
        If active_parts.Count = 0 Then
            no_species = True
        Else
            no_species = False
        End If

        'can't proceed if there isn't a part number for at least one species
        If no_species Then
            form_result = iLogicForm.ShowGlobal("epicor_13launch_species").Result
            iLogicVb.RunExternalRule("dummy.vb")

            If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
                Return
            End If
        End If
    Loop While no_species

    MultiValue.List("PartNumberToUse") = active_parts

    Dim part_selected As Boolean = False
    Dim pn As String = ""
    Do
        form_result = iLogicForm.ShowGlobal("epicor_15part_select", FormMode.Modal).Result

        If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
            Return
        End If

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
    form_result = iLogicForm.ShowGlobal("epicor_20part_properties", FormMode.Modal).Result

    If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
        Return
    End If

    form_result = check_logic(app)
    
    If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
        Return
    End If

    iLogicVb.RunExternalRule("30set_props.vb")

    'if part export fails, abort - this will usually mean the part is already
    'in the DB and so the straight add operation failed
    Dim dmt_obj As New DMT()
    Dim ret_value = PartExport.part_export(app, inv_params, dmt_obj)
    If ret_value = 0 Then
        PartRevExport.part_rev_export(app, inv_params, dmt_obj)
        PartPlantExport.part_plant_export(app, inv_params, dmt_obj)
    ElseIf ret_value = -1 Then
        MsgBox("Error: DMT timed out. Aborting...")
    Else
        MsgBox("Warning: this part is already present in Epicor. Aborting...")
    End If

    'TODO: display message box about DMT state - maybe last 3 lines of logfile
End Sub

Function check_logic(ByRef app As Inventor.Application) As FormResult
    'set a few parameters depending on data entered in first form
    Dim inv_params As UserParameters = InventorOps.get_param_set(app)
    Dim is_part_purchased As Boolean

    If StrComp(inv_params.Item("PartType").Value, "P") = 0
        is_part_purchased = True
    Else
        is_part_purchased = False
    End If

    inv_params.Item("IsPartPurchased").Value = is_part_purchased

    Dim form_result As FormResult = FormResult.OK

    Dim fails_validation As Boolean = False
    Dim required_params As New Dictionary(Of String, String) From _
            {{"PartType", "Part Type"}, {"ProdCode", "Group"}, _
             {"ClassID", "Class"}}

    'do the actual validation - there aren't many keyboard-entered fields, so
    'the most important thing to check for is that values were selected from
    'the dropdowns
    Do
        Dim error_log As String = ""
        For Each kvp As KeyValuePair(Of String, String) in required_params
            If StrComp(inv_params.Item(kvp.Key).Value, "") = 0 Then
                error_log = error_log & System.Environment.Newline & _
                            "- Select a value for " & kvp.Value
                fails_validation = True
            End If
        Next

        'set the flag to false if no errors were detected in THIS iteration
        If StrComp(error_log, "") = 0 Then
            fails_validation = False
        End If

        If fails_validation Then
            MsgBox("Please correct the following problems with the part info:" & _
                   error_log)
            form_result = iLogicForm.ShowGlobal("epicor_20part_properties", FormMode.Modal).Result
            iLogicVb.RunExternalRule("dummy.vb")

            If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
                Exit Do
            End If
        End If
    Loop While fails_validation

    Return form_result
End Function
