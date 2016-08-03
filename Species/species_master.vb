AddVbFile "inventor_common.vb"      'InventorOps.get_param_set
AddVbFile "species_list.vb"         'Species.species_list
AddVbFile "species_common.vb"       'SpeciesOps.part_pattern and mat_pattern

Imports System.Text.RegularExpressions

Sub Main()
    Dim form_result As FormResult = FormResult.OK
    'call the rules/open the forms in order to setup the iProperties properly
    iLogicVb.RunExternalRule("10species_parameters.vb")
    form_result = iLogicForm.ShowGlobal("species_20select", FormMode.Modal).Result

    If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
        Return
    End If

    'enable materials only if the species is selected AND we're working on a part doc
    Dim inv_app As Inventor.Application = ThisApplication
    Dim inv_doc As Document = inv_app.ActiveEditDocument
    Dim inv_params As UserParameters = InventorOps.get_param_set(inv_app)
    Dim is_part_doc As Boolean = TypeOf inv_doc Is PartDocument
    For Each s As String in Species.species_list
        Dim subst As String = Replace(s, "-", "4")

        If Not String.Equals(s, "Hardware") Then
            inv_params.Item("FlagMat" & subst).Value = is_part_doc AndAlso _
                    inv_params.Item("Flag" & subst).Value
        End If
    Next

    form_result = iLogicForm.ShowGlobal("species_30partnum", FormMode.Modal).Result
    If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
        Return
    End If

    form_result = validate_species()
    If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
        Return
    End If

    iLogicVb.RunExternalRule("40species_iproperties.vb")

    MsgBox("Part number iProperties successfully updated.")
End Sub

'validate the parameters for enabled species, and relaunch the form if necessary
Function validate_species() As FormResult
    Dim app As Application = ThisApplication
    Dim inv_params As UserParameters = InventorOps.get_param_set(app)

    Dim inv_doc As Document = app.ActiveEditDocument
    Dim is_part_doc As Boolean = TypeOf inv_doc Is PartDocument

    Dim part_pattern As String = "^" & SpeciesOps.part_pattern & "$"
    Dim part_regex As New Regex(part_pattern)
    Dim mat_pattern As String = "^" & SpeciesOps.mat_pattern & "$"
    Dim mat_regex As New Regex(mat_pattern)

    Dim pn_list As New List(Of String)()

    Dim fails_validation As Boolean = False

    Dim form_result As FormResult = FormResult.OK

    Do          'loop first for initial validation, check condition later
        Dim needs_reentry As String = ""
        pn_list.Clear()

        For Each s As String In Species.species_list
            Dim subst As String = Replace(s, "-", "4")
            Dim flag_value = inv_params.Item("Flag" & subst).Value
            Dim mat_flag_value = inv_params.Item("FlagMat" & subst).Value
            
            Dim materials_only As Boolean = inv_params.Item("MaterialsOnly").Value
            Dim is_intermediate_part As Boolean = inv_params.Item("IntermediatePart").Value

            If flag_value Then
                Dim part_param As Parameter = inv_params.Item("Part" & subst)
                Dim part_value As String = part_param.Value

                If materials_only Then
                    'skip checking Part fields since only materials are relevant
                ElseIf pn_list.Contains(part_value.ToUpper()) Then
                    needs_reentry = needs_reentry & System.Environment.Newline & _
                                    "- " & "Part (" & s & ") - duplicate part number"
                    fails_validation = True
                ElseIf Not is_intermediate_part Then
                    'if it's not an intermediate part, skip the regex check (since the part number
                    ' will be specified by the customer
                ElseIf StrComp(part_value, "") = 0 OrElse Not part_regex.IsMatch(part_value) Then
                    needs_reentry = needs_reentry & System.Environment.Newline & _
                                    "- " & "Part (" & s & ")"
                    fails_validation = True
                End If

                pn_list.Add(part_value.ToUpper())

                'Hardware parts and Assemblies don't have materials associated, so skip those
                ElseIf StrComp(s, "Hardware") <> 0 AndAlso mat_flag_value Then
                    Dim mat_param As Parameter = inv_params.Item("Mat" & subst)
                    Dim mat_value As String = mat_param.Value

                    If StrComp(mat_value, "") = 0 OrElse Not mat_regex.IsMatch(mat_value) Then
                        needs_reentry = needs_reentry & System.Environment.Newline & _
                                        "- " & "Material (" & s & ")"
                        fails_validation = True
                    ElseIf pn_list.Contains(mat_value) Then
                        needs_reentry = needs_reentry & System.Environment.Newline & _
                                        "- " & "Material (" & s & ") - duplicate part number"
                        fails_validation = True
                    End If

                    pn_list.Add(mat_value)
                End If
            End If
        Next

        If StrComp(needs_reentry, "") = 0 Then
            fails_validation = False
        End If

        If fails_validation Then
            MsgBox("Some entered part numbers don't fit the formatting requirements:" & _
                   needs_reentry)
            form_result = iLogicForm.ShowGlobal("species_30partnum", FormMode.Modal).Result
            iLogicVb.RunExternalRule("dummy.vb")

            If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
                Exit Do
            End If
        End If
    Loop While fails_validation

    Return form_result
End Function
