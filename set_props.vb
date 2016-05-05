'set iProperties with values the user has defined in a form
'note: these values will mostly be the IDs the Epicor DMT is expecting rather
'      than the human-readable strings
Sub Main()
    'list of parameters that need to be converted to iProperties
    Dim params = New String() {"PartType", "ProdCode", "ClassID", "UsePartRev"}

    'mappings for human-readable values (i.e. in the dropdown boxes) -> keys
    'only necessary for ProdCode and ClassID
    Dim ProdCodeMap As List(Of KeyValuePair(Of String, String)) =
        New List(Of KeyValuePair(Of String, String))
    ProdCodeMap.Add(New KeyValuePair(Of String, String)("FSC Assemblies", "FSC-ASBL"))
    ProdCodeMap.Add(New KeyValuePair(Of String, String)("FSC Components", "FSC-COMP"))
    ProdCodeMap.Add(New KeyValuePair(Of String, String)("NCA Assemblies", "NCA-ASBL"))
    ProdCodeMap.Add(New KeyValuePair(Of String, String)("NCA Components", "NCA-COMP"))
    ProdCodeMap.Add(New KeyValuePair(Of String, String)("Purchases", "PURCHASE"))

    Dim ClassIDMap As List(Of KeyValuePair(Of String, String)) =
        New List(Of KeyValuePair(Of String, String))
    ClassIDMap.Add(New KeyValuePair(Of String, String)("Assemblies we sell", "ASBL"))
    ClassIDMap.Add(New KeyValuePair(Of String, String)("Box Materials", "BOX"))
    ClassIDMap.Add(New KeyValuePair(Of String, String)("Components we sell", "COMP"))
    ClassIDMap.Add(New KeyValuePair(Of String, String)("Finished Components for kits", "FKIT"))
    ClassIDMap.Add(New KeyValuePair(Of String, String)("Finish Materials", "FNSH"))
    ClassIDMap.Add(New KeyValuePair(Of String, String)("FSC Lumber", "FSC"))
    ClassIDMap.Add(New KeyValuePair(Of String, String)("Finished Components on shelf", "FSHL"))
    ClassIDMap.Add(New KeyValuePair(Of String, String)("IT Supplies", "IT"))
    ClassIDMap.Add(New KeyValuePair(Of String, String)("Lumber" "LUMB"))
    ClassIDMap.Add(New KeyValuePair(Of String, String)("Office Supplies", "OFF"))
    ClassIDMap.Add(New KeyValuePair(Of String, String)("Other Materials", "OTHR"))
    ClassIDMap.Add(New KeyValuePair(Of String, String)("Tooling", "TOOL"))
End Sub

'TODO: call this after the parameters are set by the user to update the
'      corresponding iProperty value, in case the parameter value can't be
'      directly exported
Sub addProp(ByVal n As String)
    'define custom property collection
    oCustomPropertySet = ThisDoc.Document.PropertySets.Item("Inventor User Defined Properties")

    Try
        'set property value
        oProp = oCustomPropertySet.Item(n)
    Catch
        ' Assume error means not found so create it
        oCustomPropertySet.Add("", n)
    End Try

    'TODO: change UniqueFxName to the name of your User Defined Parameter
    iProperties.Value("Custom", n) = UniqueFxName
    'processes update when rule is run so save doesn't have to occur to see change
    iLogicVb.UpdateWhenDone = True
End Sub
