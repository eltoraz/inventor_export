'Set this rule to run when on the iTrigger button
trigger = iTrigger0

'Call the other rules in order
iLogicVb.RunExternalRule("01multi_value.vb")
iLogicForm.ShowGlobal("02part_properties", FormMode.Modal)
iLogicVb.RunExternalRule("03set_props.vb")
iLogicVb.RunExternalRule("04part_export.vb")
