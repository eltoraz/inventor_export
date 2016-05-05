'create parameters with a restricted set of accepted values for import by
'Epicor DMT (the actual IDs the tool needs are set in set_props.vb)
Sub Main()
    Dim propertyName1 As String = "PartType"
    Dim propertyName2 As String = "ProdCode"
    Dim propertyName3 As String = "ClassID"
    Dim propertyName4 As String = "UsePartRev"

    createParam(propertyName1, UnitsTypeEnum.kTextUnits)
    createParam(propertyName2, UnitsTypeEnum.kTextUnits)
    createParam(propertyName3, UnitsTypeEnum.kTextUnits)
    createParam(propertyName4, UnitsTypeEnum.kBooleanUnits)

    MultiValue.SetList(propertyName1, "M", "P")
    MultiValue.SetList(propertyName2, "FSC Assemblies", "FSC Components", "NCA Assemblies", "NCA Components", "Purchases")
    MultiValue.SetList(propertyName3, "Assemblies we sell", "Box Materials", "Components we sell", "Finished Components for kits", "Finish Materials", "FSC Lumber", "Finished Components on shelf", "IT Supplies", "Lumber", "Office Supplies", "Other Materials", "Tooling")
End Sub

Sub createParam(ByVal n As String, ByVal paramType As UnitsTypeEnum)
    dim oDoc As Document

    oDoc = ThisApplication.ActiveDocument
    Dim oParams As UserParameters = oDoc.Parameters.UserParameters

    Dim defaultValue
    If paramType = UnitsTypeEnum.kTextUnits Then
        defaultValue = ""
    ElseIf paramType = UnitsTypeEnum.kBooleanUnits Then
        defaultValue = True
    End If

    Dim TestParam As UserParameter = oParams.AddByValue(n, defaultValue, paramType)
    'TestParam.ExposedAsProperty = True
End Sub
