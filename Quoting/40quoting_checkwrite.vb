AddVbFile "parameters.vb"           'ParameterOps.get_param_set
AddVbFile "quoting_common.vb"       'QuotingOps.sheet_name

Sub Main()
    Dim inv_app As Inventor.Application = ThisApplication
    Dim inv_params As UserParameters = ParameterOps.get_param_set(inv_app)
    Dim quoting_spreadsheet As String = inv_params.Item("QuotingSpreadsheet").Value
    Dim pn As String = SpeciesOps.unpack_pn(inv_params.Item("PartNumberToUse").Value).Item1.ToUpper()

    Try
        GoExcel.Open(quoting_spreadsheet, QuotingOps.sheet_name)
    Catch ex As Exception
        MsgBox("Cannot open file to verify data was written. Error: " & ex.Message)
        Return
    End Try

    Dim data_start_row As Integer = 4
    Dim start_search As Integer = data_start_row
    Dim max_search As Integer = data_start_row + 100
    Dim working_row As Integer
    GoExcel.TitleRow = 2
    GoExcel.FindRowStart = data_start_row
    working_row = GoExcel.FindRow(quoting_spreadsheet, QuotingOps.sheet_name, "Stock Name", "=", pn)

    Dim failed = False
    Dim fail_msg As String = ""
    If working_row = -1 Then
        failed = True
    Else
        If GoExcel.CellValue("A" & working_row) <> pn OrElse _
                GoExcel.CellValue("C" & working_row) <> inv_params.Item("FinishedThickness").Value OrElse _
                GoExcel.CellValue("D" & working_row) <> inv_params.Item("Width").Value OrElse _
                GoExcel.CellValue("E" & working_row) <> inv_params.Item("Length").Value OrElse _
                GoExcel.CellValue("F" & working_row) <> inv_params.Item("WidthSpec").Value OrElse _
                GoExcel.CellValue("G" & working_row) <> inv_params.Item("LengthSpec").Value OrElse _
                GoExcel.CellValue("H" & working_row) <> inv_params.Item("QtyPerUnit").Value OrElse _
                GoExcel.CellValue("I" & working_row) <> inv_params.Item("NestedQty").Value OrElse _
                GoExcel.CellValue("J" & working_row) <> inv_params.Item("SandingSpec").Value OrElse _
                GoExcel.CellValue("K" & working_row) <> inv_params.Item("GrainDirection").Value OrElse _
                GoExcel.CellValue("L" & working_row) <> inv_params.Item("CertifiedClass").Value OrElse _
                GoExcel.CellValue("M" & working_row) <> inv_params.Item("WoodSpecies").Value OrElse _
                GoExcel.CellValue("N" & working_row) <> inv_params.Item("GlueUpSpec").Value OrElse _
                GoExcel.CellValue("O" & working_row) <> inv_params.Item("ColorSpec").Value OrElse _
                GoExcel.CellValue("P" & working_row) <> inv_params.Item("GradeSpec").Value OrElse _
                GoExcel.CellValue("Q" & working_row) <> inv_params.Item("CustomSpec").Value OrElse _
                GoExcel.CellValue("R" & working_row) <> inv_params.Item("CustomDetails").Value Then
            failed = True
        End If
    End If

    If failed Then
        MsgBox("Failed to update the spreadsheet. Check that no one else " & _
               "has it open or contact your system administrator.")
    End If

    GoExcel.Save
    GoExcel.Close
End Sub
