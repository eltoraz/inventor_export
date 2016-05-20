Dim dmt_log As String = ""

'Call the other rules in order
iLogicVb.RunExternalRule("01multi_value.vb")
iLogicForm.ShowGlobal("02part_properties", FormMode.Modal)
iLogicVb.RunExternalRule("03set_props.vb")
iLogicVb.RunExternalRule("04part_export.vb")
iLogicVb.RunExternalRule("05partrev_export.vb")

MsgBox(DMT.dmt_log)
