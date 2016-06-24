Sub Main()
    'call the rules/open the forms in order to setup the iProperties properly
    iLogicVb.RunExternalRule("10species_parameters.vb")
    iLogicForm.ShowGlobal("species_20select", FormMode.Modal)
    iLogicForm.ShowGlobal("species_30partnum", FormMode.Modal)
    iLogicVb.RunExternalRule("40species_validation.vb")
    iLogicVb.RunExternalRule("50species_iproperties.vb")

    MsgBox("Part number iProperties successfully updated.")
End Sub
