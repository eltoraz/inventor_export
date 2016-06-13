AddVbFile "dmt.vb"
AddVbFile "40part_export.vb"
AddVbFile "50partrev_export.vb"
AddVbFile "60partplant_export.vb"

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
Dim dmt_obj As New DMT()
Dim ret_value = PartExport.part_export(ThisApplication, dmt_obj)
If ret_value = 0 Then
    PartRevExport.part_rev_export(ThisApplication, dmt_obj)
    PartPlantExport.part_plant_export(ThisApplication, dmt_obj)
ElseIf ret_value = -1 Then
    MsgBox("Error: DMT timed out. Aborting...")
Else
    MsgBox("Warning: this part is already present in Epicor. Aborting...")
End If

'TODO: display message box about DMT state - maybe last 3 lines of logfile
