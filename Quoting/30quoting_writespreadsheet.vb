AddVbFile "inventor_common.vb"      'InventorOps.get_param_set
AddVbFile "quoting_common.vb"       'QuotingOps.sheet_name
AddVbFile "species_common.vb"       'SpeciesOps.unpack_pn

Sub Main()
    Dim app As Application = ThisApplication
    Dim inv_params As UserParameters = InventorOps.get_param_set(app)
    Dim quoting_spreadsheet As String = inv_params.Item("QuotingSpreadsheet").Value
    Dim pn As String = SpeciesOps.unpack_pn(inv_params.Item("PartNumberToUse").Value).Item1.ToUpper()

    Try
        GoExcel.Open(quoting_spreadsheet, QuotingOps.sheet_name)
    Catch ex As Exception
        MsgBox("Cannot open file. Error: " & ex.Message)
        Return
    End Try

    'find an open row, or this part's existing entry
    'NOTE: the spreadsheets this is designed for have no blank rows in between data items:
    '      starting at row 4, it's all data, so a blank row indicates a free spot for new entries
    '      (also, GoExcel.FindRow stops searching at the first empty cell in one of its query columns)
    Dim data_start_row As Integer = 4
    Dim start_search As Integer = data_start_row
    Dim max_search As Integer = data_start_row + 100
    Dim working_row As Integer

    'Yes/No message box returns 6 (Yes) or 7 (No) (X/Cancel is disabled)
    Dim msg_result As Integer = 7

    'first see if the part already has an entry in the spreadsheet
    GoExcel.TitleRow = 2

    GoExcel.FindRowStart = data_start_row
    working_row = GoExcel.FindRow(quoting_spreadsheet, QuotingOps.sheet_name, "Stock Name", "=", pn)

    If working_row <> -1 Then
        'if the part is already in the spreadsheet, give the user the option to continue or abort
        msg_result = MsgBox("This material already has an entry in the spreadsheet. " & _
                            "Would you like to continue to update it?", MsgBoxStyle.YesNo)
        If msg_result = 7 Then Return
    Else
        'entry not present, so look for a blank row
        Do
            For i = start_search To max_search
                If String.IsNullOrEmpty(GoExcel.CellValue("A" & count)) Then
                    working_row = i
                    Exit For
                End If
            Next

            'no blank row found in search range, so prompt user to continue or abort
            If working_row = -1 Then
                msg_result = MsgBox("The spreadsheet already has at least " & _
                                max_search-data_start_row & " entries. Continue?", _
                                MsgBoxStyle.YesNo)
                If msg_result = 7 Then Return
            End If
            start_search = max_search
            max_search += 100
        Loop While working_row = -1
    End If

    'write the quoting fields to their respective cells
    GoExcel.CellValue("A" & working_row) = pn

    GoExcel.CellValue("C" & working_row) = inv_params.Item("FinishedThickness").Value
    GoExcel.CellValue("D" & working_row) = inv_params.Item("Width").Value
    GoExcel.CellValue("E" & working_row) = inv_params.Item("Length").Value
    GoExcel.CellValue("F" & working_row) = inv_params.Item("WidthSpec").Value
    GoExcel.CellValue("G" & working_row) = inv_params.Item("LengthSpec").Value

    GoExcel.CellValue("H" & working_row) = inv_params.Item("QtyPerUnit").Value
    GoExcel.CellValue("I" & working_row) = inv_params.Item("NestedQty").Value
    GoExcel.CellValue("J" & working_row) = inv_params.Item("SandingSpec").Value
    GoExcel.CellValue("K" & working_row) = inv_params.Item("GrainDirection").Value
    GoExcel.CellValue("L" & working_row) = inv_params.Item("CertifiedClass").Value
    GoExcel.CellValue("M" & working_row) = inv_params.Item("WoodSpecies").Value
    GoExcel.CellValue("N" & working_row) = inv_params.Item("GlueUpSpec").Value
    GoExcel.CellValue("O" & working_row) = inv_params.Item("ColorSpec").Value
    GoExcel.CellValue("P" & working_row) = inv_params.Item("GradeSpec").Value
    GoExcel.CellValue("Q" & working_row) = inv_params.Item("CustomSpec").Value
    GoExcel.CellValue("R" & working_row) = inv_params.Item("CustomDetails").Value

    'description string is built from the other fields via an Excel formula,
    ' so don't need to construct/write it here (though the formula will be used
    ' in the Epicor module for raw material desc)

    GoExcel.Save
    GoExcel.Close
End Sub
