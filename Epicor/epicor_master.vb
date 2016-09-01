AddVbFile "dmt.vb"                      'DMT
AddVbFile "epicor_common.vb"            'EpicorOps.fetch_list_values
AddVbFile "40part_export.vb"            'PartExport.part_export
AddVbFile "50partrev_export.vb"         'PartRevExport.part_rev_export
AddVbFile "60partplant_export.vb"       'PartPlantExport.part_plant_export
AddVbFile "species_common.vb"           'SpeciesOps.select_active_part
AddVbFile "quoting_common.vb"           'QuotingOps.generate_desc
AddVbFile "parameters.vb"               'ParameterOps.get_param_set, species_list
AddVbFile "inventor_common.vb"          'InventorOps.format_csv_field

'master rule to export part to Epicor inventory, calling the others
Sub Main()
    'populate the PartNumberToUse param multi-value with the activated part numbers
    Dim app As Inventor.Application = ThisApplication
    Dim inv_params As UserParameters = ParameterOps.get_param_set(app)

    Dim form_result As FormResult = FormResult.OK

    'setup the suite's parameters
    iLogicVb.RunExternalRule("10multi_value.vb")

    'select the part to work on (placed in "PartNumberToUse" Inventor User Parameter)
    Dim parts_and_mats = "MP"
    If inv_params.Item("MaterialsOnly").Value Then parts_and_mats = "P"
    form_result = SpeciesOps.select_active_part(app, inv_params, ParameterOps.species_list, _
                                                iLogicForm, iLogicVb, MultiValue, parts_and_mats)
    If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
        Return
    End If

    'set some parameters based on the type of the selected part
    Dim design_props As PropertySet = app.ActiveDocument.PropertySets.Item("Design Tracking Properties")
    Dim part As Tuple(Of String, String, String) = SpeciesOps.unpack_pn(inv_params.Item("PartNumberToUse").Value)
    Dim part_species As String = part.Item3
    Dim is_part As Boolean = inv_params.Item("ActiveIsPart").Value
    If is_part Then
        inv_params.Item("Description").Value = design_props.Item("Description").Value
    Else
        Try
            inv_params.Item("Description").Value = QuotingOps.generate_desc(part_species, inv_params)
            inv_params.Item("UsePartRev").Value = False
            inv_params.Item("TrackSerialNum").Value = False
        Catch ex As Exception
            MsgBox("Warning: the fields for this raw material haven't been setup. " & _
                   "Try running the Quoting Spreadsheet export first.")
            Return
        End Try
    End If

    'set multi-value lists for ProdCode & ClassID based on the selected part type
    MultiValue.List("ProdCode") = EpicorOps.fetch_list_values("ProdCode.csv", _
                                                              DMT.dmt_working_path, _
                                                              part.Item2)
    MultiValue.List("ClassID") = EpicorOps.fetch_list_values("ClassID.csv", _
                                                             DMT.dmt_working_path, _
                                                             part.Item2)

    'Call the other rules in order
    form_result = iLogicForm.ShowGlobal("epicor_20part_properties", FormMode.Modal).Result
    If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
        Return
    End If

    form_result = check_logic(app)
    If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
        Return
    End If

    iLogicVb.RunExternalRule("30set_props.vb")

    'if the flag for the current part/mat shows it's already been exported, abort
    Dim flag As String = "Exported"
    If is_part Then
        flag = flag & "Part"
    Else
        flag = flag & "Mat"
    End If
    flag = flag & Replace(part_species, "-", "4")
    If inv_params.Item(flag).Value Then
        MsgBox("Part/species combination """ & part.Item1 & "/" & part_species & _
               """ has already been exported into Epicor from this document. Aborting...")
        Return
    End If

    'if any part of the export fails, abort - this will usually mean
    ' the part is already in the DB and so the straight add operation failed
    Dim dmt_obj As New DMT()
    Dim ret_value = PartExport.part_export(app, inv_params, dmt_obj)
    If ret_value <> 0 Then
        dmt_obj.check_errors(ret_value, "Part")
        Return
    End If

    ret_value = PartRevExport.part_rev_export(app, inv_params, dmt_obj)
    If ret_value <> 0 Then
        dmt_obj.check_errors(ret_value, "Part Revision")
        Return
    End If

    ret_value = PartPlantExport.part_plant_export(app, inv_params, dmt_obj)
    If ret_value <> 0 Then
        dmt_obj.check_errors(ret_value, "Part Plant")
        Return
    End If

    inv_params.Item(flag).Value = True
    MsgBox("DMT has successfully imported part " & part.Item1 & " into Epicor.")
End Sub

'validate the form logic, and return a form result
' (if reentry required) that lets the user abort
Function check_logic(ByRef app As Inventor.Application) As FormResult
    Dim inv_doc As Document = app.ActiveDocument
    Dim inv_params As UserParameters = ParameterOps.get_param_set(app)
    Dim design_props As PropertySet = inv_doc.PropertySets.Item("Design Tracking Properties")
    Dim summary_props As PropertySet = inv_doc.PropertySets.Item("Inventor Summary Information")

    Dim form_result As FormResult = FormResult.OK

    Dim fails_validation As Boolean = False
    Dim required_params As New Dictionary(Of String, String) From _
            {{"ProdCode", "Group"}, {"ClassID", "Class"}}

    'do the actual validation - there aren't many keyboard-entered fields, so
    'the most important thing to check for is that values were selected from
    'the dropdowns
    Do
        Dim error_log As String = ""
        Dim description As String = inv_params.Item("Description").Value

        '1/1/1601 = Win32 epoch = null date for Inventor date fields
        Dim appr_date, null_date As Date
        appr_date = design_props.Item("Engr Date Approved").Value
        null_date = #1/1/1601#

        If String.IsNullOrEmpty(description) Then
            error_log = error_log & System.Environment.NewLine & _
                        "- Enter a description"
            fails_validation = True
        End If

        'validate revision number as two homogenous characters
        Dim rev_regex As New System.Text.RegularExpressions.Regex("^(\d{2}|[A-Za-z]{2})$")
        Dim rev_match As System.Text.RegularExpressions.Match = rev_regex.Match(summary_props.Item("Revision Number").Value)
        If Not rev_match.Success Then
            error_log = error_log & System.Environment.NewLine & _
                        "- Enter a 2-digit or 2-letter revision ID"
            fails_validation = True
        End If

        For Each kvp As KeyValuePair(Of String, String) In required_params
            If String.IsNullOrEmpty(inv_params.Item(kvp.Key).Value) Then
                error_log = error_log & System.Environment.NewLine & _
                            "- Select a value for " & kvp.Value
                fails_validation = True
            End If
        Next

        If appr_date = null_date Then
            error_log = error_log & System.Environment.NewLine & _
                        "- Select an approval date"
            fails_validation = True
        End If

        'set the flag to false if no errors were detected in THIS iteration
        If String.IsNullOrEmpty(error_log) Then
            fails_validation = False
        End If

        If fails_validation Then
            MsgBox("Please correct the following problems with the part info:" & _
                   error_log)
            form_result = iLogicForm.ShowGlobal("epicor_20part_properties", FormMode.Modal).Result

            If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
                Exit Do
            End If
        End If
    Loop While fails_validation

    Return form_result
End Function