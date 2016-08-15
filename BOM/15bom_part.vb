AddVbFile "dmt.vb"                      'DMT
AddVbFile "parameters.vb"               'ParameterOps.get_param_set
AddVbFile "species_common.vb"           'SpeciesOps.select_active_part
AddVbFile "bom_common.vb"               'BomOps.bom_fields

Sub Main()
    Dim inv_app As Inventor.Application = ThisApplication
    Dim inv_doc As Document = inv_app.ActiveDocument
    Dim inv_params As UserParameters = ParameterOps.get_param_set(inv_app)

    Dim summary_props As PropertySet = inv_doc.PropertySets.Item("Inventor Summary Information")

    Dim selected_part As Tuple(Of String, String, String) = _
            SpeciesOps.unpack_pn(inv_params.Item("PartNumberToUse").Value)

    Dim PartNum, RevisionNum, MtlPartNum As String
    Dim MtlSeq As Integer = BomOps.MtlSeqStart 

    'get the part number of the associated material
    Dim part_species As String = selected_part.Item3
    Dim subst As String = Replace(part_species, "-", "4")
    MtlPartNum = inv_params.Item("Mat" & subst).Value

    PartNum = selected_part.Item1
    RevisionNum = summary_props.Item("Revision Number").Value

    'naively assume that the quantity of materials needed for the part only
    ' depends on the nested quantity from quoting (eg, ignoring cases where
    ' multiple *different* parts use a *single* raw material)
    Dim QtyPer As Double = 1 / inv_params.Item("NestedQty").Value

    Dim data As String
    data = BomOps.bom_values("Company")                         'Company name (constant)
    data = data & "," & PartNum
    data = data & "," & RevisionNum
    data = data & "," & MtlSeq
    data = data & "," & MtlPartNum
    data = data & "," & QtyPer
    data = data & "," & BomOps.bom_values("Plant")              'Plant (constant)
    data = data & "," & BomOps.bom_values("ECOGroupID")         'ECO Group (constant)

    Dim dmt_obj As New DMT()
    Dim file_name As String
    file_name = dmt_obj.write_csv("Bill_Of_Materials.csv", BomOps.bom_fields, data)

    Dim ret_code As Integer = dmt_obj.dmt_import("Bill Of Materials", file_name, False)
End Sub
