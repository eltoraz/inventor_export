' <IsStraightVb>True</IsStraightVb>
Imports Inventor

Public Class EpicorOps
    'master list of parameters created for Epicor module
    Public Shared param_list As New Dictionary(Of String, UnitsTypeEnum) From _
            {{"PartType", UnitsTypeEnum.kTextUnits}, _
             {"ProdCode", UnitsTypeEnum.kTextUnits}, _
             {"ClassID", UnitsTypeEnum.kTextUnits}, _
             {"UsePartRev", UnitsTypeEnum.kBooleanUnits}, _
             {"MfgComment", UnitsTypeEnum.kTextUnits}, _
             {"PurComment", UnitsTypeEnum.kTextUnits}, _
             {"TrackSerialNum", UnitsTypeEnum.kBooleanUnits}, _
             {"RevDescription", UnitsTypeEnum.kTextUnits}, _
             {"PartNumberToUse", UnitsTypeEnum.kTextUnits}}

    'enclose field in quotes, and escape quotes already in the field
    Public Shared Function format_csv_field(ByVal s As String) As String
        Dim s2 As String = Replace(s, """", """""")
        Return """" & s2 & """"
    End Function
End Class