AddVbFile "parameters.vb"           'ParameterOps.create_all_params
AddVbFile "species_common.vb"       'SpeciesOps.unpack_pn

Imports Inventor

Sub Main()
    Dim inv_app As Application = ThisApplication

    ParameterOps.create_all_params(inv_app)
    Dim inv_params As UserParameters = ParameterOps.get_param_set(inv_app)

    For Each kvp As KeyValuePair(Of String, Tuple(Of UnitsTypeEnum, ArrayList)) In ParameterOps.quoting_params
        Dim valid_values As ArrayList = kvp.Value.Item2
        If valid_values.Count > 0 Then
            MultiValue.List(kvp.Key) = valid_values
        End If
    Next

    'special cases
    'populate the Wood Species with the value from part selection
    Dim part_entry As String = inv_params.Item("PartNumberToUse").Value
    Dim part_unpacked As Tuple(Of String, String, String) = SpeciesOps.unpack_pn(part_entry)
    Dim wood_species As String = part_unpacked.Item3
    inv_params.Item("WoodSpecies").Value = Replace(wood_species, "-", "_").ToUpper()

    'set the Color Spec options correctly for the selected species
    If ParameterOps.quoting_color_specs.ContainsKey(wood_species) Then
        Dim color_spec_options As New ArrayList(ParameterOps.quoting_params("ColorSpec").Item2)
        color_spec_options.AddRange(ParameterOps.quoting_color_specs(wood_species))
        MultiValue.List("ColorSpec") = color_spec_options
    End If

    'if the Color Spec has been set for this species, retrieve it from the parameter
    Dim prev_color_spec As Parameter = inv_params.Item("ColorSpec" & Replace(wood_species, "-", "4"))
    If Not String.IsNullOrEmpty(prev_color_spec.Value) Then
        inv_params.Item("ColorSpec").Value = prev_color_spec.Value
    End If
End Sub
