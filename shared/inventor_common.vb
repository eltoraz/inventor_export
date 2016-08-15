' <IsStraightVb>True</IsStraightVb>
'Inventor parameter/iProperty manipulation functions
Imports Inventor

Public Class InventorOps
    'update iProperty `n` with value `param_val`, creating it if it doesn't exist
    Public Shared Sub update_prop(ByVal n As String, ByVal param_val As Object, _
                                  ByRef app As Inventor.Application)
        'get the custom property collection
        Dim inv_doc As Document = app.ActiveDocument
        Dim inv_custom_props As PropertySet 
        inv_custom_props = inv_doc.PropertySets.Item("Inventor User Defined Properties")

        ' Attempt to get existing custom property
        On Error Resume Next
        Dim prop
        prop = inv_custom_props.Item(n)
        If Err.Number <> 0 Then
            'Failed to get the property, which means it doesn't already exist,
            'so we'll create it
            inv_custom_props.Add(param_val, n)
        Else
            'got the property so update the value
            prop.Value = param_val
        End If
    End Sub

    'enclose field in quotes, and escape quotes already in the field
    Public Shared Function format_csv_field(ByVal s As String) As String
        Dim s2 As String = Replace(s, """", """""")
        Return """" & s2 & """"
    End Function
End Class
