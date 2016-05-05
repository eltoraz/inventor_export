'set iProperties with values the user has defined in a form
'note: these values will mostly be the IDs the Epicor DMT is expecting rather
'      than the human-readable strings
Sub Main()
    'list of parameters that need to be converted to iProperties
    Dim params = New String() {"PartType", "ProdCode", "ClassID", "UsePartRev"}

    'mappings for human-readable values (i.e. in the dropdown boxes) -> keys
    'only necessary for ProdCode and ClassID
    Dim ProdCodeMap As New Dictionary(Of String, String)
    ProdCodeMap.Add("FSC Assemblies", "FSC-ASBL")
    ProdCodeMap.Add("FSC Components", "FSC-COMP")
    ProdCodeMap.Add("NCA Assemblies", "NCA-ASBL")
    ProdCodeMap.Add("NCA Components", "NCA-COMP")
    ProdCodeMap.Add("Purchases", "PURCHASE")

    Dim ClassIDMap As New Dictionary(Of String, String)
    ClassIDMap.Add("Assemblies we sell", "ASBL")
    ClassIDMap.Add("Box Materials", "BOX")
    ClassIDMap.Add("Components we sell", "COMP")
    ClassIDMap.Add("Finished Components for kits", "FKIT")
    ClassIDMap.Add("Finish Materials", "FNSH")
    ClassIDMap.Add("FSC Lumber", "FSC")
    ClassIDMap.Add("Finished Components on shelf", "FSHL")
    ClassIDMap.Add("IT Supplies", "IT")
    ClassIDMap.Add("Lumber", "LUMB")
    ClassIDMap.Add("Office Supplies", "OFF")
    ClassIDMap.Add("Other Materials", "OTHR")
    ClassIDMap.Add("Tooling", "TOOL")

    For Each item As String in params
        Dim paramVal
    Next
End Sub

Sub updateProp(ByVal n As String, ByVal paramVal As Variant)
    'get the custom property collection
    Dim invDoc As Document = ThisApplication.ActiveDocument
    Dim invCustomPropertySet = invDoc.PropertySets.Item("Inventor User Defined Properties")

    ' Attempt to get existing custom property
    On Error Resume Next
    Dim invProp As Property
    Set invProperty = invCustomPropertySet.Item(n)
    If Err.Number <> 0 Then
        'Failed to get the property, which means it doesn't already exist,
        'so we'll create it
        Call invCustomPropertySet.Add(paramVal, n)
    Else
        'got the property so update the value
        invProperty.value = paramVal
    End If

    'processes update when rule is run so save doesn't have to occur to see change
    'iLogicVb.UpdateWhenDone = True
End Sub
