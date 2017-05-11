' <IsStraightVb>True</IsStraightVb>
Imports System.Windows.Forms
Imports Inventor
Imports Autodesk.iLogic.Interfaces

Public Module QuotingOps
    Public starting_path As String = "N:\CompanyResources\Quoting\"
    Public sheet_name As String = "BBN Quote Sheet"

    'display a dialog to select the quoting spreadsheet to use for the current part
    'set the QuotingSpreadsheet parameter to the path & filename, and test opening it
    'return DialogResult.OK if successful, DialogResult.Cancel if the user cancels
    ' (or the file can't be opened)
    Public Function pick_spreadsheet(ByRef inv_params As UserParameters, _
                                     ByRef GoExcel As IGoExcel) As DialogResult
        'open the quoting spreadsheet
        'using VB-native dialog instead of Inventor since navigating network drives is easier
        Dim file_picker As New OpenFileDialog()
        file_picker.InitialDirectory = starting_path
        file_picker.Title = "Select Quoting spreadsheet to use..."
        file_picker.Filter = "Microsoft Excel Spreadsheets (*.xls, *.xlsx)|*.xls;*.xlsx|All Files (*.*)|*.*"
        file_picker.FilterIndex = 1

        Dim current_param_value As String = inv_params.Item("QuotingSpreadsheet").Value
        If Not String.IsNullOrEmpty(current_param_value) Then
            file_picker.FileName = System.IO.Path.GetFileName(current_param_value)
            file_picker.InitialDirectory = System.IO.Path.GetDirectoryName(current_param_value)
        End If

        Dim dialog_result As DialogResult = file_picker.ShowDialog()
        If dialog_result = DialogResult.OK Then
            inv_params.Item("QuotingSpreadsheet").Value = file_picker.FileName
            Try
                GoExcel.Open(file_picker.FileName, sheet_name)
                'even though we're not making changes, need to save before closing
                ' for Inventor to properly release the file
                GoExcel.Save
                GoExcel.Close
            Catch ex As Exception
                MsgBox("Cannot open file. Error: " & ex.Message)
                Return DialogResult.Cancel
            End Try
        End If

        Return dialog_result
    End Function

    'return a string containing the description for the given raw material
    'note: this uses parameters populated in the quoting module for the specified species
    Public Function generate_desc(ByVal species As String, _
                                  ByRef inv_params As UserParameters) As String
        Dim desc As String = ""
        Dim subst As String = Replace(species, "-", "4")

        'desc1 (in old format)
        desc = desc & String.Format("{0:#.0000}", inv_params.Item("FinishedThickness").Value) & "T_X_"
        desc = desc & String.Format("{0:#.0000}", inv_params.Item("Width").Value) & "W_X_"
        desc = desc & String.Format("{0:#.0000}", inv_params.Item("Length").Value) & "L-"
        desc = desc & inv_params.Item("WidthSpec").Value & "-"
        desc = desc & inv_params.Item("LengthSpec").Value & "-"
        desc = desc & inv_params.Item("SandingSpec").Value & "-"
        desc = desc & inv_params.Item("GrainDirection").Value & "-"

        'desc2 (in old format)
        desc = desc & inv_params.Item("CertifiedClass").Value & "-"
        desc = desc & inv_params.Item("WoodSpecies").Value & "-"
        desc = desc & inv_params.Item("GlueUpSpec").Value & "-"
        desc = desc & inv_params.Item("ColorSpec" & subst).Value & "-"
        desc = desc & inv_params.Item("GradeSpec").Value & "-MLD"

        If Not String.IsNullOrEmpty(inv_params.Item("CustomSpec").Value) Then
            desc = desc & "_" & inv_params.Item("CustomSpec").Value
        End If
        If Not String.IsNullOrEmpty(inv_params.Item("CustomDetails").Value) Then
            desc = desc & "_" & inv_params.Item("CustomDetails").Value
        End If

        Return desc
    End Function

    'validate quoting form data, prompting for reentry of required fields
    Function validate_quoting(ByVal require_spreadsheet As Boolean, _
                              ByRef inv_params As UserParameters, _
                              ByRef app As Inventor.Application, _
                              ByRef iLogicForm As IiLogicForm) As FormResult
        Dim fails_validation As Boolean = False
        Dim required_text_fields As New Dictionary(Of String, String) From _
                {{"WidthSpec", "Width Spec"}, _
                {"LengthSpec", "Length Spec"}, {"SandingSpec", "Sanding Spec"}, _
                {"GrainDirection", "Grain Direction"}, {"CertifiedClass", "Certified Classification"}, _
                {"GlueUpSpec", "Glue up or solid stock"}, {"GradeSpec", "Grade Spec"}}
        Dim required_num_fields As New Dictionary(Of String, String) From _
                {{"FinishedThickness", "Finished Thickness"}, _
                {"Length", "Length"}, {"QtyPerUnit", "Qty Per Unit"}, _
                {"NestedQty", "Nested Qty"}}

        Dim form_result As FormResult = FormResult.OK

        If Not String.Equals(inv_params.Item("SandingSpec").Value, "SDIA_150") Then
            required_num_fields.Add("Width", "Width")
        End If

        'validation loop - most values are selected from multi-value lists, and the
        ' default values tend to be "" or 0.0
        Do
            Dim error_log = ""

            'different log message for spreadsheet path
            If require_spreadsheet AndAlso _
                    String.IsNullOrEmpty(inv_params.Item("QuotingSpreadsheet").Value) Then

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
                                "- Set a nonzero value for " & kvp.Value
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
End Module
