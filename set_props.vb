'set iProperties with values the user has defined in a form
'note: these values will mostly be the IDs the Epicor DMT is expecting rather
'      than the human-readable strings
Sub Main()
    'list of parameters that need to be converted to iProperties
    Dim params = New String() {"PartType", "ProdCode", "ClassID", "UsePartRev", "MfgComment", "PurComment", "TrackSerialNum"}

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

    For Each i As String in params
        Dim invDoc As Document = ThisApplication.ActiveDocument
        Dim invParams As UserParameters = invDoc.Parameters.UserParameters

        'if Epicor requires a short ID, convert the human-readable value via
        'the appropriate mapping (see above)
        'required for: ProdCode, ClasID
        Dim param As Parameter = invParams.Item(i)
        Dim paramValue = param.Value
        If StrComp(i, "ProdCode") = 0 Then
            paramValue = ProdCodeMap(paramValue)
        Else If StrComp(i, "ClassID") = 0 Then
            paramValue = ClassIDMap(paramValue)
        Else If StrComp(i, "MfgComment") = 0 Then
            'note: Epicor MfgComment field supports up to 16000 chars)
            paramValue = Left(paramValue, 16000)
        Else If StrComp(i, "PurComment") = 0 Then
            'note: Epicor PurComment field supports up to 16000 chars)
            paramValue = Left(paramValue, 16000)
        End If

        updateProp(i, paramValue)

        invDoc.Update
    Next
End Sub

Sub updateProp(ByVal n As String, ByVal paramVal As Object)
    'get the custom property collection
    Dim invDoc As Document = ThisApplication.ActiveDocument
    Dim invCustomPropertySet As PropertySet 
    invCustomPropertySet = invDoc.PropertySets.Item("Inventor User Defined Properties")

    ' Attempt to get existing custom property
    On Error Resume Next
    Dim invProp
    invProp = invCustomPropertySet.Item(n)
    If Err.Number <> 0 Then
        'Failed to get the property, which means it doesn't already exist,
        'so we'll create it
        invCustomPropertySet.Add(paramVal, n)
    Else
        'got the property so update the value
        invProp.value = paramVal
    End If
End Sub
