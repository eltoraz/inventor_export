AddVbFile "inventor_common.vb"      'InventorOps.get_param_set
AddVbFile "species_list.vb"         'Species.species_list
AddVbFile "species_common.vb"       'SpeciesOps.select_active_part
AddVbFile "quoting_common.vb"       'QuotingOps.starting_path

Imports System.Windows.Forms

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim inv_params As UserParameters = InventorOps.get_param_set(app)   

    Dim form_result As FormResult = FormResult.OK

    'select the part to work with
    form_result = SpeciesOps.select_active_part(app, inv_params, Species.species_list, _
                                                iLogicForm, iLogicVb, MultiValue)
    If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
        Return
    End If

    'setup the parameters for the form
    iLogicVb.RunExternalRule("10quoting_parameters.vb")

    form_result = iLogicForm.ShowGlobal("quoting_20field_entry", FormMode.Modal).Result
    If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
        Return
    End If

    'validate form data
    form_result = validate_quoting(app)
    If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
        Return
    End If

    'write data to spreadsheet
    iLogicVb.RunExternalRule("30quoting_writespreadsheet.vb")
End Sub

Function validate_quoting(ByRef app As Inventor.Application) As FormResult
    'TODO: pop up a form to hand-enter value for "Molded" if "Custom" selected
    Dim form_result As FormResult = FormResult.OK
    Return form_result
End Function
