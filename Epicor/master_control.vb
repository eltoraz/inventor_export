AddVbFile "dmt.vb"
AddVbFile "40part_export.vb"

'Pull latest data from Epicor
'this data shouldn't change often, so the rule shouldn't need to be called often
'DMT.dmt_export()

'Call the other rules in order
iLogicVb.RunExternalRule("10multi_value.vb")
iLogicForm.ShowGlobal("20part_properties", FormMode.Modal)
'iLogicVb.RunExternalRule("21logic_check.vb")
iLogicVb.RunExternalRule("30set_props.vb")

'if part export fails, abort - this will usually mean the part is already
'in the DB and so the straight add operation failed
Dim ret_value = PartExport.part_export()
If ret_value = 0 Then
    iLogicVb.RunExternalRule("50partrev_export.vb")
    iLogicVb.RunExternalRule("60partplant_export.vb")
End If

'TODO: display message box about DMT state - maybe last 3 lines of logfile
