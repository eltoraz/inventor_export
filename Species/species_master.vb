'call the rules/open the forms in order to setup the iProperties properly
iLogicVb.RunExternalRule("10species_parameters.vb")
iLogicForm.ShowGlobal("20species_select", FormMode.Modal)
iLogicForm.ShowGlobal("30species_partnum", FormMode.Modal)
iLogicVb.RunExternalRule("40species_iproperties.vb")
