'validate the parameters for enabled species, and relaunch the form if necessary

Sub Main()
    Dim inv_doc As Document = ThisApplication.ActiveEditDocument
    Dim part_doc As PartDocument
    Dim assm_doc As AssemblyDocument
    Dim inv_params As UserParameters

    If TypeOf inv_doc Is PartDocument Then
        part_doc = app.ActiveEditDocument
        inv_params = part_doc.ComponentDefinition.Parameters.UserParameters
    ElseIf TypeOf inv_doc Is AssemblyDocument Then
        assm_doc = app.ActiveEditDocument
        inv_params = assm_doc.ComponentDefinition.Parameters.UserParameters
    Else
        inv_params = inv_doc.ComponentDefinition.Parameters.UserParameters
    End If

    Dim needs_reentry As NewArrayList()
    For i = 1 To inv_params.Count
        'stuff
    Next
End Sub
