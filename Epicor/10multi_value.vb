AddVbFile "parameters.vb"           'ParameterOps.create_all_params

'create parameters with a restricted set of accepted values for import by
'Epicor DMT (the actual IDs the tool needs are set in set_props.vb)
Sub Main()
    Dim inv_app As Inventor.Application = ThisApplication

    ParameterOps.create_all_params(inv_app)

    MultiValue.SetList("PartType", "M", "P")

    'contents of multi-value lists for ProdCode & ClassID depend on manufactured
    ' part vs raw material, so they need to be populated in master after part selection
End Sub
