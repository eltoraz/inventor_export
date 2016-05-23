AddVbFile "dmt.vb"

'Pull latest data from Epicor
'TODO: this shouldn't change often, so it shouldn't need to be called often
DMT.dmt.export()

'Call the other rules in order
iLogicVb.RunExternalRule("01multi_value.vb")
iLogicForm.ShowGlobal("02part_properties", FormMode.Modal)
iLogicVb.RunExternalRule("03set_props.vb")
iLogicVb.RunExternalRule("04part_export.vb")
iLogicVb.RunExternalRule("05partrev_export.vb")

'TODO: display message box about DMT state - maybe last 3 lines of logfile
