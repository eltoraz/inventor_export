AddVbFile "dmt.vb"                      'DMT

Sub Main()
    'Pull latest data from Epicor
    'this data shouldn't change often, so the rule shouldn't need to be called often
    Dim dmt_obj As New DMT()
    dmt_obj.dmt_export()
End Sub
