AddVbFile "inventor_common.vb"      'InventorOps.create_param
AddVbFile "parameters.vb"           'ParameterLists.quoting_params
AddVbFile "species_list.vb"         'Species.species_list

Sub Main()
    Dim inv_app As Inventor.Application = ThisApplication

    'create shared parameters (if they don't exist) along with this module's
    For Each kvp As KeyValuePair(Of String, UnitsTypeEnum) In ParameterLists.shared_params
        InventorOps.create_param(kvp.Key, kvp.Value, inv_app)
    Next
    For Each kvp As KeyValuePair(Of String, Tuple(Of UnitsTypeEnum, ArrayList)) In ParameterLists.quoting_params
        InventorOps.create_param(kvp.Key, kvp.Value.Item1, inv_app)
        
        Dim valid_values As ArrayList = kvp.Value.Item2
        If valid_values.Count > 0 Then
            MultiValue.List(kvp.Key) = valid_values
        End If
    Next

    'create color spec parameters for each species
    For Each s As String in Species.species_list
        Dim subst As String = Replace(s, "-", "4")
        InventorOps.create_param("ColorSpec" & subst, UnitsTypeEnum.kTextUnits, inv_app)
    Next
End Sub
