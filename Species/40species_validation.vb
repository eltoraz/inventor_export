AddVbFile "species_list.vb"         'Species.species_list

Imports System.Text.RegularExpressions

'validate the parameters for enabled species, and relaunch the form if necessary
Sub Main()
    Dim app As Application = ThisApplication
    Dim inv_doc As Document = app.ActiveEditDocument
    Dim part_doc As PartDocument
    Dim assm_doc As AssemblyDocument
    Dim inv_params As UserParameters

    If TypeOf inv_doc Is PartDocument Then
        part_doc = app.ActiveEditDocument
        inv_params = part_doc.ComponentDefinition.Parameters.UserParameters
    ElseIf TypeOf inv_doc Is AssemblyDocument Then
        assm_doc = app.ActiveEditDocument
        inv_params = assm_doc.ComponentDefinition.Parameters.UserParameters
    Else
        inv_params = inv_doc.ComponentDefinition.Parameters.UserParameters
    End If

    'regular expression to match the part number format AZ-123
    Dim partno_pattern As String = "[a-zA-Z]{2}-[0-9]{3}"
    Dim partno_regex As New Regex(partno_pattern)
    Dim fails_validation As Boolean = False

    Do          'loop first for initial validation, check condition later
        Dim needs_reentry As String = ""

        For Each s As String In Species.species_list
            Dim subst As String = Replace(s, "-", "4")
            Dim flag_param As Parameter = inv_params.Item("Flag" & subst)
            Dim flag_value = flag_param.Value

            If flag_value Then
                Dim part_param As Parameter = inv_params.Item("Part" & subst)
                Dim part_value As String = part_param.Value

                If StrComp(part_value, "") = 0 OrElse Not partno_regex.IsMatch(part_value) Then
                    needs_reentry = needs_reentry & System.Environment.Newline & _
                                    "- " & "Part (" & s & ")"
                    fails_validation = True
                End If

                If StrComp(s, "Hardware") <> 0 Then
                    Dim mat_param As Parameter = inv_params.Item("Mat" & subst)
                    Dim mat_value As String = mat_param.Value

                    If StrComp(mat_value, "") = 0 OrElse Not partno_regex.IsMatch(mat_value) Then
                        needs_reentry = needs_reentry & System.Environment.Newline & _
                                        "- " & "Material (" & s & ")"
                        fails_validation = True
                    End If
                End If
            End If
        Next

        If StrComp(needs_reentry, "") = 0 Then
            fails_validation = False
        End If

        If fails_validation Then
            MsgBox("Some entered part numbers don't fit the formatting requirements:" & _
                   needs_reentry)
            iLogicForm.ShowGlobal("30species_partnum", FormMode.Modal)
            iLogicVb.RunExternalRule("dummy.vb")
        End If
    Loop While fails_validation
End Sub
