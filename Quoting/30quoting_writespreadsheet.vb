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
    'TODO: need to account for spreadsheet schema, since GoExcel.FindRow stops
    '       searching on the first empty cell in the column
    Dim working_row As Integer
    GoExcel.TitleRow = 2
    GoExcel.FindRowStart = 4
    working_row = GoExcel.FindRow(quoting_spreadsheet, QuotingOps.sheet_name, "Stock Name", "=", pn)

    'DEBUG
    MsgBox(quoting_spreadsheet & "; " & QuotingOps.sheet_name & System.Environment.NewLine & _ 
            pn & "; " & working_row)
    If working_row = -1 Then
        working_row = 17
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
