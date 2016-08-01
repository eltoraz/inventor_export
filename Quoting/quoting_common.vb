' <IsStraightVb>True</IsStraightVb>
Imports System.Windows.Forms
Imports Inventor
Imports Autodesk.iLogic.Interfaces

Public Class QuotingOps
    Public Shared starting_path As String = "N:\CompanyResources\Quoting\"
    Public Shared sheet_name As String = "NFP Quote Sheet"

    'display a dialog to select the quoting spreadsheet to use for the current part
    'set the QuotingSpreadsheet parameter to the path & filename, and test opening it
    'return DialogResult.OK if successful, DialogResult.Cancel if the user cancels
    ' (or the file can't be opened)
    Public Shared Function pick_spreadsheet(ByRef inv_params As UserParameters, _
                                            ByRef GoExcel As IGoExcel) As DialogResult
        'open the quoting spreadsheet
        'using VB-native dialog instead of Inventor since navigating network drives is easier
        Dim file_picker As New OpenFileDialog()
        file_picker.InitialDirectory = starting_path
        file_picker.Title = "Select Quoting spreadsheet to use..."
        file_picker.Filter = "Microsoft Excel Spreadsheets (*.xls, *.xlsx)|*.xls;*.xlsx|All Files (*.*)|*.*"
        file_picker.FilterIndex = 1

        Dim dialog_result As DialogResult = file_picker.ShowDialog()
        If dialog_result = DialogResult.OK Then
            inv_params.Item("QuotingSpreadsheet").Value = file_picker.FileName
            Try
                GoExcel.Open(file_picker.FileName, sheet_name)
                GoExcel.Close()
            Catch ex As Exception
                MsgBox("Cannot open file. Error: " & ex.Message)
                Return DialogResult.Cancel
            End Try
        End If

        Return dialog_result
    End Function
End Class
