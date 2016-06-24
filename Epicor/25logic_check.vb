AddVbFile "inventor_common.vb"      'InventorOps.get_param_set

Sub Main()
    'set a few parameters depending on data entered in first form
    Dim app As Inventor.Application = ThisApplication
    Dim inv_params As UserParameters = InventorOps.get_param_set(app)
    Dim is_part_purchased As Boolean

    If StrComp(inv_params.Item("PartType").Value, "P") = 0
        is_part_purchased = True
    Else
        is_part_purchased = False
    End If

    inv_params.Item("IsPartPurchased").Value = is_part_purchased

    Dim fails_validation As Boolean = False
    Dim required_params As New Dictionary(Of String, String) From _
            {{"PartType", "Part Type"}, {"ProdCode", "Group"}, _
             {"ClassID", "Class"}}

    'do the actual validation - there aren't many keyboard-entered fields, so
    'the most important thing to check for is that values were selected from
    'the dropdowns
    Do
        Dim error_log As String = ""
        For Each kvp As KeyValuePair(Of String, String) in required_params
            If StrComp(inv_params.Item(kvp.Key).Value, "") = 0 Then
                error_log = error_log & System.Environment.Newline & _
                            "- Select a value for " & kvp.Value
                fails_validation = True
            End If
        Next

        'set the flag to false if no errors were detected in THIS iteration
        If StrComp(error_log, "") = 0 Then
            fails_validation = False
        End If

        If fails_validation Then
            MsgBox("Please correct the following problems with the part info:" & _
                   error_log)
            iLogicForm.ShowGlobal("epicor_20part_properties", FormMode.Modal)
            iLogicVb.RunExternalRule("dummy.vb")
        End If
    Loop While fails_validation
End Sub
