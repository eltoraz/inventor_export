AddVbFile "inventor_common.vb"      'InventorOps.get_param_set
AddVbFile "quoting_common.vb"       'QuotingOps.pick_spreadsheet

Imports System.Windows.Forms
Imports Inventor

Sub Main()
    Dim inv_params As UserParameters = InventorOps.get_param_set(ThisApplication)
    QuotingOps.pick_spreadsheet(inv_params, GoExcel)
End Sub
