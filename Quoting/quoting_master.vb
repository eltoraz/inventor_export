AddVbFile "inventor_common.vb"      'InventorOps.get_param_set
AddVbFile "species_list.vb"         'Species.species_list
AddVbFile "species_common.vb"       'SpeciesOps.select_active_part

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim inv_params As UserParameters = InventorOps.get_param_set(app)   

    Dim form_result As FormResult = FormResult.OK

    'setup the parameters this module needs
    iLogicVb.RunExternalRule("10quoting_parameters.vb")

    'select the part to work with
    form_result = SpeciesOps.select_active_part(app, inv_params, Species.species_list, _
                                                iLogicForm, iLogicVb, MultiValue)
    If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
        Return
    End If
End Sub
