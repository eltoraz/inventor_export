AddVbFile "inventor_common.vb"      'InventorOps.create_param
AddVbFile "parameters.vb"           'ParameterLists.quoting_params
AddVbFile "species_list.vb"         'Species.species_list
AddVbFile "species_common.vb"       'SpeciesOps.unpack_pn

Imports Inventor

Sub Main()
    Dim inv_app As Application = ThisApplication
    Dim inv_params As UserParameters = InventorOps.get_param_set(inv_app)   

    'run through shared parameters to make sure they're set up
    For Each kvp As KeyValuePair(Of String, UnitsTypeEnum) in ParameterLists.shared_params
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

    'special cases
    'populate the Wood Species with the value from part selection
    Dim part_entry As String = inv_params.Item("PartNumberToUse").Value
    Dim part_unpacked As Tuple(Of String, String, String) = SpeciesOps.unpack_pn(part_entry)
    Dim wood_species As String = part_unpacked.Item3
    inv_params.Item("WoodSpecies").Value = Replace(wood_species, "-", "_").ToUpper()

    'set the Color Spec options correctly for the selected species
    If ParameterLists.quoting_color_specs.ContainsKey(wood_species) Then
        Dim color_spec_options As New ArrayList(ParameterLists.quoting_params("ColorSpec").Item2)
        color_spec_options.AddRange(ParameterLists.quoting_color_specs(wood_species))
        MultiValue.List("ColorSpec") = color_spec_options
    End If

    'if the Color Spec has been set for this species, retrieve it from the parameter
    Dim prev_color_spec As Parameter = inv_params.Item("ColorSpec" & Replace(wood_species, "-", "4"))
    If Not String.IsNullOrEmpty(prev_color_spec.Value) Then
        inv_params.Item("ColorSpec").Value = prev_color_spec.Value
    End If
End Sub
