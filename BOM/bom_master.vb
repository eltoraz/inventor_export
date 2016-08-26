AddVbFile "15bom_part.vb"               'PartBOMExport.part_bom_export
AddVbFile "16bom_asm.vb"                'AssmBOMExport.assm_bom_export
AddVbFile "species_common.vb"           'SpeciesOps.select_active_part
AddVbFile "parameters.vb"               'ParameterOps.create_all_params, get_param_set, species_list
AddVbFile "dmt.vb"                      'DMT

Sub Main()
    Dim inv_app As Inventor.Application = ThisApplication
    Dim inv_doc As Document = inv_app.ActiveEditDocument
    Dim inv_params As UserParameters = ParameterOps.get_param_set(inv_app)

    Dim form_result As FormResult = FormResult.OK

    'create missing parameters
    ParameterOps.create_all_params(inv_app)

    'select the part to export the BOM for (only manufactured parts, which will
    ' pull in their associated raw materials/component parts anyway)
    form_result = SpeciesOps.select_active_part(app, inv_params, ParameterOps.species_list, _
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

    Dim dmt_obj As New DMT()

    Dim bom_return_code As Integer
    'BOM export procedure for parts and assemblies is different
    If TypeOf inv_doc Is PartDocument Then
        bom_return_code = PartBOMExport.part_bom_export(inv_app, inv_params, dmt_obj)
    ElseIf TypeOf inv_doc Is AssemblyDocument Then
        bom_return_code = AssmBOMExport.assm_bom_export(inv_app, inv_params, ThisBOM, dmt_obj)
    Else
        MsgBox("Error: MOM can only be exported from a Part or Assembly. Aborting...")
        Return
    End If
End Sub
