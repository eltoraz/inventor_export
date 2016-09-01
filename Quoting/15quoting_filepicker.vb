AddVbFile "parameters.vb"           'ParameterOps.get_param_set
AddVbFile "quoting_common.vb"       'QuotingOps.pick_spreadsheet

Imports System.Windows.Forms
Imports Inventor

'wrapper rule for pick_spreadsheet to hook to a form button
Sub Main()
    Dim inv_params As UserParameters = ParameterOps.get_param_set(ThisApplication)
    QuotingOps.pick_spreadsheet(inv_params, GoExcel)
End Sub
