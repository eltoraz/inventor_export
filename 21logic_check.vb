'set a few parameters depending on data entered in first form
Dim inv_doc As Document = ThisApplication.ActiveDocument
Dim inv_params As UserParameters = inv_doc.Parameters.UserParameters
Dim is_part_purchased As Boolean

If StrComp(inv_params.Item("PartType"), "P") = 0
    is_part_purchased = True
Else
    is_part_purchased = False
End If

inv_params.Item("IsPartPurchased").Value = is_part_purchased
