'create parameters with a restricted set of accepted values for import by
'Epicor DMT (the actual IDs the tool needs are set in set_props.vb)
Sub Main()
    Dim params As New Dictionary(Of String, UnitsTypeEnum)
    params.Add("PartType", UnitsTypeEnum.kTextUnits)
    params.Add("ProdCode", UnitsTypeEnum.kTextUnits)
    params.Add("ClassID", UnitsTypeEnum.kTextUnits)
    params.Add("UsePartRev", UnitsTypeEnum.kBooleanUnits)
    params.Add("MfgComment", UnitsTypeEnum.kTextUnits)
    params.Add("PurComment", UnitsTypeEnum.kTextUnits)
    params.Add("TrackSerialNum", UnitsTypeEnum.kBooleanUnits)

    For Each kvp As KeyValuePair(Of String, UnitsTypeEnum) in params
        createParam(kvp.Key, kvp.Value)
    Next

    MultiValue.SetList("PartType", "M", "P")
    MultiValue.SetList("ProdCode", "FSC Assemblies", "FSC Components", "NCA Assemblies", "NCA Components", "Purchases")
    MultiValue.SetList("ClassID", "Assemblies we sell", "Box Materials", "Components we sell", "Finished Components for kits", "Finish Materials", "FSC Lumber", "Finished Components on shelf", "IT Supplies", "Lumber", "Office Supplies", "Other Materials", "Tooling")
End Sub

Sub createParam(ByVal n As String, ByVal paramType As UnitsTypeEnum)
    dim invDoc As Document = ThisApplication.ActiveDocument

    Dim invParams As UserParameters = invDoc.Parameters.UserParameters

    Dim defaultValue
    If paramType = UnitsTypeEnum.kTextUnits Then
        defaultValue = ""
    ElseIf paramType = UnitsTypeEnum.kBooleanUnits Then
        defaultValue = True
    End If

    Dim TestParam As UserParameter = invParams.AddByValue(n, defaultValue, paramType)

    invDoc.Update
End Sub
