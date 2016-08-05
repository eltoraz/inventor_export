' <IsStraightVb>True</IsStraightVb>

Public Class BomOps
    Public Shared bom_fields As String = "Company,PartNum,RevisionNum,MtlSeq,MtlPartNum,QtyPer,Plant,ECOGroupID"

    Public shared MtlSeqStart As Integer = 10
    Public Shared bom_values As New Dictionary(Of String, String) From _
            {{"Company", "BBN"}, {"Plant", "MfgSys"}, {"ECOGroupID", "DMT"}}
End Class
