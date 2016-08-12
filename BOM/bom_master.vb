AddVbFile "species_common.vb"           'SpeciesOps.select_active_part
AddVbFile "species_list.vb"             'Species.species_list
AddVbFile "inventor_common.vb"          'InventorOps.get_param_set

Sub Main()
    Dim inv_app As Inventor.Application = ThisApplication
    Dim inv_doc As Document = inv_app.ActiveEditDocument
    Dim inv_params As UserParameters = InventorOps.get_param_set(inv_app)

    Dim form_result As FormResult = FormResult.OK

    'select the part to export the BOM for (only manufactured parts, which will
    ' pull in their associated raw materials/component parts anyway)
    form_result = SpeciesOps.select_active_part(app, inv_params, Species.species_list, _
                                                iLogicForm, iLogicVb, MultiValue, "M")
    If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
        Return
    End If

    'check `Exported` flags for part and associated material, warn if both aren't True
    'TODO: consider other cases (eg, abort in cases where only materials are defined)
    'TODO: validate other values (eg, revision number, though this should be caught
    '      in Epicor export)
    Dim selected_part As Tuple(Of String, String, String) = _
            SpeciesOps.unpack_pn(inv_params.Item("PartNumberToUse").Value)
    Dim part_species As String = selected_part.Item3
    Dim subst As String = Replace(part_species, "-", "4")

    Dim warn_before_continue As Boolean = False
    Dim warn_message As String = "The part or material for " & part_species & _
            " may not have been exported to Epicor inventory. Please run Epicor export on:"

    Try
        Dim exported_part As Boolean = inv_params.Item("ExportedPart" & subst).Value
        Dim exported_mat As Boolean = inv_params.Item("ExportedMat" & subst).Value

        If Not exported_part Then
            warn_before_continue = True
            warn_message = warn_message & System.Environment.NewLine & _
                           "- Part (" & part_species & ")"
        End If
        If Not exported_mat Then
            warn_before_continue = True
            warn_message = warn_message & System.Environment.NewLine & _
                           "- Material (" & part_species & ")"
        End If
    Catch ex As Exception
        warn_before_continue = True
        warn_message = "Couldn't determine whether the part and material have been" & _
                       " exported into Epicor inventory."
    End Try
    warn_message = warn_message & System.Environment.NewLine & System.Environment.NewLine & _
                   "Continue anyway?"

    'prompt to continue if we're unsure whether the part/mat are in inventory
    'Yes/No message box returns 6 (Yes) or 7 (No) (X/Cancel is disabled)
    Dim msg_result As Integer = 6
    If warn_before_continue Then msg_result = MsgBox(warn_message, MsgBoxStyle.YesNo)
    If msg_result = 7 Then Return

    'BOM export procedure for parts and assemblies is different
    If TypeOf inv_doc Is PartDocument Then
        iLogicVb.RunExternalRule("15bom_part.vb")
    ElseIf TypeOf inv_doc Is AssemblyDocument Then
        iLogicVb.RunExternalRule("16bom_asm.vb")
    Else
        MsgBox("Error: MOM can only be exported from a Part or Assembly. Aborting...")
        Return
    End If
End Sub
