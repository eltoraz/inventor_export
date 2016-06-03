AddVbFile "dmt.vb"

'Pull latest data from Epicor
'this data shouldn't change often, so the rule shouldn't need to be called often
'DMT.dmt_export()

'Call the other rules in order
iLogicVb.RunExternalRule("10multi_value.vb")
iLogicForm.ShowGlobal("20part_properties", FormMode.Modal)
'iLogicVb.RunExternalRule("21logic_check.vb")
iLogicVb.RunExternalRule("30set_props.vb")
iLogicVb.RunExternalRule("40part_export.vb")
iLogicVb.RunExternalRule("50partrev_export.vb")
iLogicVb.RunExternalRule("60partplant_export.vb")

'TODO: display message box about DMT state - maybe last 3 lines of logfile
