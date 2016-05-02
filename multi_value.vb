Sub Main()
    'change Custom iProperty Name to desired Custom iProperty Name
    Dim propertyName1 As String = "PartType"
    Dim propertyName2 As String = "Group"
    Dim propertyName3 As String = "ClassID"
    Dim propertyName4 As String = "UsePartRev"

    createParam(propertyName1)
    createParam(propertyName2)
    createParam(propertyName3)
    createParamBool(propertyName4)

    MultiValue.SetList(propertyName1, "M", "P")
    MultiValue.SetList(propertyName2, "FSC Assemblies", "FSC Componenets", "NCA Assemblies", "NCA Components", "Purchases")
    MultiValue.SetList(propertyName3, "Assemblies we sell", "Box Materials", "Components we sell", "Finish Materials", "Finished Components for kits", "Finished Components on shelf", "FSC Lumber", "Lumber", "Other Materials", "Tooling")

    'addProp(propertyName1)
    'addProp(propertyName2)
    'addProp(propertyName3)
End Sub

Sub createParam(ByVal n As String)
    dim oDoc As Document

    oDoc = ThisApplication.ActiveDocument
    Dim oParams As UserParameters = oDoc.Parameters.UserParameters

    Dim TestParam As UserParameter = oParams.AddByValue(n, "", UnitsTypeEnum.kTextUnits)
    TestParam.ExposedAsProperty = True
End Sub

Sub createParamBool(ByVal n As String)
    dim oDoc As Document

    oDoc = ThisApplication.ActiveDocument
    Dim oParams As UserParameters = oDoc.Parameters.UserParameters

    Dim TestParam As UserParameter = oParams.AddByValue(n, True, UnitsTypeEnum.kBooleanUnits)
    TestParam.ExposedAsProperty = True
End Sub

Sub addProp(ByVal n As String)
    'define custom property collection
    oCustomPropertySet = ThisDoc.Document.PropertySets.Item("Inventor User Defined Properties")

    Try
        'set property value
        oProp = oCustomPropertySet.Item(propertyName1)
    Catch
        ' Assume error means not found so create it
        oCustomPropertySet.Add("", propertyName1)
    End Try

    'set custom property value; Change Custom iProperty Name to desired Custom iProperty Name;
    'change UniqueFxName to the name of your User Defined Parameter
    iProperties.Value("Custom", "PartType") = UniqueFxName
    'processes update when rule is run so save doesn't have to occur to see change
    iLogicVb.UpdateWhenDone = True
End Sub
