' <IsStraightVb>True</IsStraightVb>
Imports Inventor

Public Class SpeciesOps
    Public Shared Function select_active_part(ByRef app As Inventor.Application) As FormResult
        'select the part we'll be working with here
        Dim active_parts As New ArrayList()
        Dim no_species As Boolean = False

        Dim form_result As FormResult = FormResult.OK

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
                        Dim part_entry As String = part_value & " - Part (" & s & ")"
                        active_parts.Add(part_entry)

                        If StrComp(s, "Hardware") <> 0 Then
                            Dim mat_param As Parameter = inv_params.Item("Mat" & subst)
                            Dim mat_value As String = mat_param.Value
                            Dim mat_entry As String = mat_value & " - Material (" & s & ")"
                            active_parts.Add(mat_entry)
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
                form_result = iLogicForm.ShowGlobal("epicor_13launch_species", FormMode.Modal).Result

                If form_result = FormResult.None Then
                    Return form_result
                End If
            End If
        Loop While no_species

        MultiValue.List("PartNumberToUse") = active_parts

        Dim part_selected As Boolean = False
        Dim pn As String = ""
        Do
            form_result = iLogicForm.ShowGlobal("epicor_15part_select", FormMode.Modal).Result

            If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
                Return form_result
            End If

            pn = inv_params.Item("PartNumberToUse").Value
            If StrComp(pn, "") <> 0 Then
                part_selected = True
            Else
                MsgBox("Please select a part to continue with the Epicor export.")
                iLogicVb.RunExternalRule("dummy.vb")
            End If
        Loop While Not part_selected

        Return form_result
    End Function
End Class
