AddVbFile "dmt.vb"                  'DMT.dmt_working_path
AddVbFile "inventor_common.vb"      'InventorOps.update_prop
AddVbFile "parameters.vb"           'ParameterOps.get_param_set
AddVbFile "epicor_common.vb"        'EpicorOps.fetch_list_mappings

'set iProperties with values the user has defined in a form (epicor_20)
'note: these values will mostly be the IDs the Epicor DMT is expecting rather
'      than the human-readable strings that will remain in the Parameters
Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim inv_doc As Document = app.ActiveDocument
    Dim inv_params As UserParameters = ParameterOps.get_param_set(app)

    'mappings for human-readable values (i.e. in the dropdown boxes) -> keys
    ' (only necessary for ProdCode and ClassID)
    Dim ProdCodeMap As Dictionary(Of String, String) = _
                EpicorOps.fetch_list_mappings("ProdCode.csv", DMT.dmt_working_path)
    Dim ClassIDMap As Dictionary(Of String, String) = _
                EpicorOps.fetch_list_mappings("ClassID.csv", DMT.dmt_working_path)

    'update description separately since it's in a different property set AND
    ' the iProperty only gets changed if we're working with an actual part
    ' rather than a raw material
    Dim design_props As PropertySet = app.ActiveDocument.PropertySets.Item("Design Tracking Properties")
    Dim is_part As Boolean = inv_params.Item("ActiveIsPart").Value
    
    'store the user-entered description in the "official" property only for parts
    ' since material descriptions can be regenerated on the fly
    If is_part Then
        design_props.Item("Description").Value = inv_params.Item("Description").Value
    End If

    inv_doc.Update

    For Each kvp As KeyValuePair(Of String, UnitsTypeEnum) in ParameterOps.epicor_params
        Dim param As Parameter = inv_params.Item(kvp.Key)
        Dim param_name As String = param.Name
        Dim param_value = param.Value

        'if Epicor requires a short ID, convert the human-readable value via
        'the appropriate mapping (see above)
        'required for: ProdCode, ClasID
        If String.Equals(param_name, "ProdCode") Then
            param_value = ProdCodeMap(param_value)
        Else If String.Equals(param_name, "ClassID") Then
            param_value = ClassIDMap(param_value)
        'truncate comment fields to 16000 chars (Epicor max)
        Else If String.Equals(param_name, "MfgComment") Then
            param_value = Left(param_value, 16000)
        Else If String.Equals(param_name, "PurComment") Then
            param_value = Left(param_value, 16000)
        Else If String.Equals(param_name, "RevDescription") Then
            param_value = Left(param_value, 16000)
        End If

        InventorOps.update_prop(param_name, param_value, app)

        inv_doc.Update
    Next
End Sub