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
    Dim bom_obj As BOM = comp_def.BOM
    bom_obj.PartsOnlyViewEnabled = True

    Dim bom_view As BOMView = bom_obj.BOMViews.Item("Parts Only")
    Dim bom_row As BOMRow

    Dim summary_props As PropertySet = inv_doc.PropertySets.Item("Inventor Summary Information")

    Dim selected_part As Tuple(Of String, String, String) = _
            SpeciesOps.unpack_pn(inv_params.Item("PartNumberToUse").Value)

    Dim PartNum, RevisionNum, MtlPartNum As String
    Dim MtlSeq As Integer = BomOps.MtlSeqStart 

    Dim part_species As String = selected_part.Item3
    Dim subst As String = Replace(part_species, "-", "4")

    PartNum = selected_part.Item1
    RevisionNum = summary_props.Item("Revison Number").Value

    Dim data As String = ""
    For Each bom_row In bom_view.BOMRows
        Dim child_comp_def As ComponentDefinition
        child_comp_def = bom_row.ComponentDefinitions.Item(1)

        Dim custom_props As PropertySet
        custom_props = child_comp_def.Document.PropertySets.Item("Inventor User Defined Properties")
        Try
            MtlPartNum = custom_props.Item("Part (" & part_species & ")").Value
        Catch e As Exception
            Dim child_filename As String = child_comp_def.Document.FullDocumentName
            MsgBox("The part number is not defined for the specified species for child part " & _
                   child_filename & ". Please run BBN Species Setup on that part, save the " & _
                   "document, and rerun this BOM export.")
            Return
        End Try

        Dim QtyPer As Integer = ThisBOM.CalculateQuantity("Model Data", unique_part)

        data = data & BomOps.bom_values("Company")                  'Company name (constant)
        data = data & "," & PartNum
        data = data & "," & RevisionNum
        data = data & "," & MtlSeq
        data = data & "," & MtlPartNum
        data = data & "," & QtyPer
        data = data & "," & BomOps.bom_values("Plant")              'Plant (constant)
        data = data & "," & BomOps.bom_values("ECOGroupID")         'ECO Group (constant)
        data = data & System.Environment.NewLine
    Next

    Dim dmt_obj As New DMT()
    Dim file_name As String
    file_name = dmt_obj.write_csv("Bill_Of_Materials.csv", BomOps.bom_fields, data)

    Dim ret_code As Integer = dmt_obj.dmt_import("Bill of Materials", file_name, False)
End Sub
