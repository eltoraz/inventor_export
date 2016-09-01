' <IsStraightVb>True</IsStraightVb>

'common properties used in the BOM rules
Public Module BomOps
    Public bom_fields As String = "Company,PartNum,RevisionNum,MtlSeq,MtlPartNum,QtyPer,Plant,ECOGroupID"

    'each material in the bill of materials needs a unique material
    ' sequence number; parts just have one since they have only one
    ' material, and assemblies need to increment this for each
    ' subsequent unique material
    Public MtlSeqStart As Integer = 10

    'centrally define static fields
    Public bom_values As New Dictionary(Of String, String) From _
            {{"Company", "BBN"}, {"Plant", "MfgSys"}, {"ECOGroupID", "DMT"}}
End Module
