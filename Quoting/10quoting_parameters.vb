AddVbFile "inventor_common.vb"      'InventorOps.create_param
AddVbFile "quoting_common.vb"       'QuotingOps.param_list

Sub Main()
    Dim inv_app As Inventor.Application = ThisApplication
    For Each kvp As KeyValuePair(Of String, UnitsTypeEnum) in QuotingOps.param_list
        InventorOps.create_param(kvp.Key, kvp.Value.Item1, inv_app)
    Next
End Sub
