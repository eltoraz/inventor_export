AddVbFile "inventor_common.vb"      'InventorOps.get_param_set
AddVbFile "species_list.vb"         'Species.species_list
AddVbFile "species_common.vb"       'SpeciesOps.select_active_part
AddVbFile "quoting_common.vb"       'QuotingOps.starting_path

Imports System.Windows.Forms
Imports Inventor

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

'validate quoting form data, prompting for reentry of required fields
Function validate_quoting(ByRef app As Inventor.Application) As FormResult
    Dim inv_params As UserParameters = InventorOps.get_param_set(app)

    Dim fails_validation As Boolean = False
    Dim required_text_fields As New Dictionary(Of String, String) From _
            {{"WidthSpec", "Width Spec"}, _
             {"LengthSpec", "Length Spec"}, {"SandingSpec", "Sanding Spec"}, _
             {"GrainDirection", "Grain Direction"}, {"CertifiedClass", "Certified Classification"}, _
             {"GlueUpSpec", "Glue up or solid stock"}, {"GradeSpec", "Grade Spec"}, _
             {"Molded", "Molded"}}
    Dim required_num_fields As New Dictionary(Of String, String) From _
            {{"FinishedThickness", "Finished Thickness"}, {"Width", "Width"}, _
             {"Length", "Length"}, {"QtyPerUnit", "Qty Per Unit"}, _
             {"NestedQty", "Nested Qty"}}

    Dim form_result As FormResult = FormResult.OK

    'validation loop - most values are selected from multi-value lists, and the
    ' default values tend to be "" or 0.0
    Do
        Dim error_log = ""

        'pop up a form to hand-enter value for "Molded" if "Custom" selected
        Dim molded_param As Parameter = inv_params.Item("Molded")
        If String.Equals(molded_param.Value, "Custom") Then
            Dim molded_val As String = InputBox("Enter custom molding specification." & _
                            System.Environment.NewLine & _
                            "(Clear the input box to abort spreadsheet generation.)", _
                            "Custom molding specification")
            If String.IsNullOrEmpty(molded_val) Then
                form_result = FormResult.Cancel
                Exit Do
            Else
                molded_param.Value = molded_val
            End If
        End If

        'different log message for spreadsheet path
        If String.IsNullOrEmpty(inv_params.Item("QuotingSpreadsheet").Value) Then
            error_log = error_log & System.Environment.NewLine & _
                        "- Select a spreadsheet to use for quoting"
            fails_validation = True
        End If

        'check required text parameters
        For Each kvp As KeyValuePair(Of String, String) In required_text_fields
            If String.IsNullOrEmpty(inv_params.Item(kvp.Key).Value) Then
                error_log = error_log & System.Environment.NewLine & _
                            "- Set a value for " & kvp.Value
                fails_validation = True
            End If
        Next

        'check required numeric parameters
        For Each kvp As KeyValuePair(Of String, String) In required_num_fields
            If inv_params.Item(kvp.Key).Value <= 0.0 Then
                error_log = error_log & System.Environment.NewLine & _
                            "- Set a value for " & kvp.Value
                fails_validation = True
            End If
        Next

        'set the flag to end the loop if validation passed on this iteration
        If String.IsNullOrEmpty(error_log) Then
            fails_validation = False
        End If

        If fails_validation Then
            MsgBox("Please correct the problems in the following fields: " & error_log)
            form_result = iLogicForm.ShowGlobal("quoting_20field_entry", FormMode.Modal).Result

            'abort if the user cancels the form
            If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
                Exit Do
            End If
        End If
    Loop While fails_validation

    Return form_result
End Function
