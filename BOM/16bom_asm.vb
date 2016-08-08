AddVbFile "dmt.vb"                      'DMT
AddVbFile "inventor_common.vb"          'InventorOps.get_param_set
AddVbFile "species_common.vb"           'SpeciesOps.select_active_part
AddVbFile "bom_common.vb"               'BomOps.bom_fields

Sub Main()
    Dim inv_app As Inventor.Application = ThisApplication
    Dim inv_doc As AssemblyDocument = inv_app.ActiveDocument
    Dim inv_params As UserParameters = InventorOps.get_param_set(inv_app)

    'BOM objects
    Dim comp_def As AssemblyComponentDefinition
    comp_def = inv_doc.ComponentDefinition
    Dim part_list As New List(Of String)()      'iLogic doesn't seem to support HashSet

    'get unique part numbers
    For Each occ in comp_def.Occurrences
        Dim occ_name As String = occ.Name
        Dim colon_pos As Integer = occ_name.IndexOf(":")
        occ_name = occ_name.Substring(0, colon_pos)
        If Not part_list.Contains(occ_name) Then part_list.Add(occ_name)
    Next

    Dim summary_props As PropertySet = inv_doc.PropertySets.Item("Inventor Summary Information")

    Dim selected_part As Tuple(Of String, String, String) = _
            SpeciesOps.unpack_pn(inv_params.Item("PartNumberToUse").Value)

    Dim PartNum, RevisionNum, MtlPartNum As String
    Dim MtlSeq As Integer = BomOps.MtlSeqStart 

    'get the part number of the associated material
    Dim part_species As String = selected_part.Item3
    Dim subst As String = Replace(part_species, "-", "4")

    Dim data As String = ""
End Sub
