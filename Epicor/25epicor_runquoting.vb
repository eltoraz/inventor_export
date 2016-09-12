AddVbFile "parameters.vb"           'ParameterOps.get_param_set
AddVbFile "quoting_common.vb"       'QuotingOps.validate_quoting
AddVbFile "species_common.vb"       'SpeciesOps.unpack_pn

Imports Inventor

Sub Main()
    Dim inv_app As Inventor.Application = ThisApplication
    Dim inv_params As UserParameters = ParameterOps.get_param_set(inv_app)

    'check whether the part is a material; abort if not
    If inv_params.Item("ActiveIsPart").Value Then
        MsgBox("The selected part is not a raw material, and so the quoting " & _
               "fields don't apply to it.")
        Return
    End If

    Dim form_result = FormResult.OK

    'no guarantee that the Quoting module has been run, so set up its
    ' parameters & multi-value lists
    iLogicVb.RunExternalRule("10quoting_parameters.vb")

    form_result = iLogicForm.ShowGlobal("quoting_20field_entry", FormMode.Modal).Result
    If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
        Return
    End If

    'validate form data (except for spreadsheet itself)
    form_result = QuotingOps.validate_quoting(False, inv_params, inv_app)
    If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
        Return
    End If

    'update generated description with newly-entered values
    Dim part_species As String = SpeciesOps.unpack_pn(inv_params.Item("PartNumberToUse").Value).Item3
    inv_params.Item("Description").Value = QuotingOps.generate_desc(part_species, inv_params)
End Sub
