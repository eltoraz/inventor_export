' <IsStraightVb>True</IsStraightVb>

Public Module BomOps
    Public bom_fields As String = "Company,PartNum,RevisionNum,MtlSeq,MtlPartNum,QtyPer,Plant,ECOGroupID"

    Public MtlSeqStart As Integer = 10
    Public bom_values As New Dictionary(Of String, String) From _
            {{"Company", "BBN"}, {"Plant", "MfgSys"}, {"ECOGroupID", "DMT"}}
End Module
