AddVbFile "04part_export.vb"
AddVbFile "05partrev_export.vb"

Dim dmt_log As String = ""

'Call the other rules in order
iLogicVb.RunExternalRule("01multi_value.vb")
iLogicForm.ShowGlobal("02part_properties", FormMode.Modal)
iLogicVb.RunExternalRule("03set_props.vb")

dmt_log = dmt_log & Part_Export.Part() & Environment.NewLine    '04part_export.vb
dmt_log = dmt_log & Part_Rev_Export.Part_Rev()                  '05part_rev_export.vb
