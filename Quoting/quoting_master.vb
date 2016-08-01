AddVbFile "inventor_common.vb"      'InventorOps.get_param_set
AddVbFile "species_list.vb"         'Species.species_list
AddVbFile "species_common.vb"       'SpeciesOps.select_active_part
AddVbFile "quoting_common.vb"       'QuotingOps.starting_path

Imports System.Windows.Forms

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

    'open the quoting spreadsheet
    'using VB-native dialog instead of Inventor since navigating network drives is easier
    Dim file_picker As New OpenFileDialog()
    file_picker.InitialDirectory = QuotingOps.starting_path
    file_picker.Title = "Select Quoting spreadsheet to use..."
    file_picker.Filter = "Microsoft Excel Spreadsheets (*.xls)|*.xls|All Files (*.*)|*.*"
    file_picker.FilterIndex = 1
    'TODO: display dialog & operate on returned filename/path

    form_result = iLogicForm.ShowGlobal("quoting_20field_entry", FormMode.Modal).Result
    If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
        Return
    End If
End Sub

Function validate_quoting(ByRef app As Inventor.Application) As FormResult
    'TODO: pop up a form to hand-enter value for "Molded" if "Custom" selected
End Function
