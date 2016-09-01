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
        desc = desc & inv_params.Item("GradeSpec").Value & "-MLD_"
        desc = desc & inv_params.Item("CustomSpec").Value & "_"
        desc = desc & inv_params.Item("CustomDetails").Value

        Return desc
    End Function
End Module
