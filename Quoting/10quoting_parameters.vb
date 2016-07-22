AddVbFile "inventor_common.vb"      'InventorOps.create_param
AddVbFile "quoting_common.vb"       'QuotingOps.param_list

Sub Main()
    Dim inv_app As Inventor.Application = ThisApplication
    For Each kvp As KeyValuePair(Of String, Tuple(Of UnitsTypeEnum, ArrayList)) in QuotingOps.param_list
        InventorOps.create_param(kvp.Key, kvp.Value.Item1, inv_app)
        
        Dim valid_values As ArrayList = kvp.Value.Item2
        If valid_values.Count > 0 Then
            MultiValue.List(kvp.Key) = valid_values
        End If
    Next
End Sub
