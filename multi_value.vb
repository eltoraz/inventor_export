Sub Main()
    'change Custom iProperty Name to desired Custom iProperty Name
    Dim propertyName1 As String = "Part Type"
    dim oDoc As Document

    oDoc = ThisDoc.Document

    'Check to see if this is a part file
    'If oDoc.DocumentType <> kPartDocumentObject Then
    '    MessageBox.Show("This rule can only be run in a " & DocType & " file - exiting rule...")
    '    Return
    'End If

    createParam("PartType")
    createParam("Group")
    createParam("ClassID")

    MultiValue.SetList("PartType", "M", "P")
    MultiValue.SetList("Group", "FSC Assemblies", "FSC Componenets", "NCA Assemblies", "NCA Components", "Purchases")
    MultiValue.SetList("ClassID", "Assemblies we sell", "Box Materials", "Components we sell", "Finish Materials", "Finished Components for kits", "Finished Components on shelf", "FSC Lumber", "Lumber", "Other Materials", "Tooling")

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
    iProperties.Value("Custom", "Part Type") = UniqueFxName
    'processes update when rule is run so save doesn't have to occur to see change
    iLogicVb.UpdateWhenDone = True
End Sub

Sub createParam(ByVal n As String)
    dim oDoc As Document
    oDoc = ThisDoc.Document

    Dim oPartCompDef As PartComponentDefinition = oDoc.ComponentDefinition
    Dim oNewParameter As Parameters = oPartCompDef.Parameters
    Dim oUParameter As UserParameters = oNewParameter.UserParameters
    Dim oParam As Parameter
    Dim oValue As String

    Try
        oValue = ""
        oParam = oNewParameter(n)
    Catch
        oParam = oUParameter.AddByValue(n, (oValue), "text")
    End Try
End Sub
