AddVbFile "parameters.vb"           'ParameterOps.get_param_set, species_list
AddVbFile "species_common.vb"       'SpeciesOps.select_active_part
AddVbFile "quoting_common.vb"       'QuotingOps.starting_path

Imports System.Windows.Forms
Imports Inventor

'master control for entering fields & populating quoting spreadsheet
Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim inv_params As UserParameters = ParameterOps.get_param_set(app)

    Dim form_result As FormResult = FormResult.OK

    'select the part to work with (only raw materials/purchased parts)
    form_result = SpeciesOps.select_active_part(app, inv_params, ParameterOps.species_list, _
                                                iLogicForm, iLogicVb, MultiValue, "P")
    If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
        Return
    End If

    'setup the parameters for the module
    iLogicVb.RunExternalRule("10quoting_parameters.vb")

    form_result = iLogicForm.ShowGlobal("quoting_20field_entry", FormMode.Modal).Result
    If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
        Return
    End If

    'validate form data
    form_result = QuotingOps.validate_quoting(True, inv_params, app, iLogicForm)
    If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
        Return
    End If

    'set species' stored color spec to the one just selected & verified
    'this allows us to persist different color specs for different parts in the same drawing
    Dim part As Tuple(Of String, String, String) = SpeciesOps.unpack_pn(inv_params.Item("PartNumberToUse").Value)
    Dim part_species As String = part.Item3
    inv_params.Item("ColorSpec" & Replace(part_species, "-", "4")).Value = _
            inv_params.Item("ColorSpec").Value

    'write data to spreadsheet
    iLogicVb.RunExternalRule("30quoting_writespreadsheet.vb")
    
    'verify data was saved
    iLogicVb.RunExternalRule("40quoting_checkwrite.vb")
End Sub
