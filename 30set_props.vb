AddVbFile "dmt.vb"

'set iProperties with values the user has defined in a form
'note: these values will mostly be the IDs the Epicor DMT is expecting rather
'      than the human-readable strings
Sub Main()
    'TODO: add purchase point and lead time once implemented
    'list of parameters that need to be converted to iProperties
    Dim params = New String() {"PartType", "ProdCode", "ClassID", "UsePartRev", "MfgComment", "PurComment", "TrackSerialNum", "RevDescription", "LeadTime", "VendorNum", "PurPoint"}

    'mappings for human-readable values (i.e. in the dropdown boxes) -> keys
    'only necessary for ProdCode and ClassID
    Dim ProdCodeMap As Dictionary(Of String, String) = fetch_list_mappings("ProdCode.csv")
    Dim ClassIDMap As Dictionary(Of String, String) = fetch_list_mappings("ClassID.csv")
    Dim VendorNumMap As Dictionary(Of String, String) = fetch_list_mappings("VendorNum.csv")

    'TODO: map approving engineers to Epicor IDs

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
        Else If StrComp(i, "VendorNum") = 0 Then
            paramValue = VendorNumMap(paramValue)
        Else If StrComp(i, "MfgComment") = 0 Then
            'note: Epicor MfgComment and PurComment fields supports up to 16000 chars,
            'and commas need to be stripped to avoid messing up the CSV
            paramValue = Replace(paramValue, ",", "")
            paramValue = Left(paramValue, 16000)
        Else If StrComp(i, "PurComment") = 0 Then
            paramValue = Replace(paramValue, ",", "")
            paramValue = Left(paramValue, 16000)
        Else If StrComp(i, "RevDescription") = 0 Then
            paramValue = Replace(paramValue, ",", "")
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

'Map the description in the parameter to the DB friendly ID expected by Epicor
Function fetch_list_mappings(ByVal f As String) As Dictionary(Of String, String)
    Dim file_name As String = DMT.dmt_working_path & "ref\" & f
    Dim mapping As New Dictionary(Of String, String)

    Using csv_reader As New FileIO.TextFieldParser(file_name)
        csv_reader.TextFieldType = FileIO.FieldType.Delimited
        csv_reader.SetDelimiters(",")

        Dim current_row As String()
        Dim first_line As Boolean = True
        While Not csv_reader.EndOfData
            Try
                current_row = csv_reader.ReadFields()
            Catch ex As FileIO.MalformedLineException
                Debug.Write("CSV contained invalid line:" & ex.Message)
            End Try

            If first_line Then
                'skip headers
                first_line = False
            Else
                mapping.Add(current_row(0), current_row(1))
            End If
        End While
    End Using

    Return mapping
End Function
