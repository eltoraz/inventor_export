' <IsStraightVb>True</IsStraightVb>
'Inventor parameter/iProperty manipulation functions
Imports Inventor

Public Class InventorOps
    'initialize parameter `n` as type `param_type`
    Public Shared Sub create_param(ByVal n As String, ByVal param_type As UnitsTypeEnum, _
                                   ByRef app As Inventor.Application)
        Dim inv_doc As Document = app.ActiveDocument
        Dim inv_params As UserParameters = get_param_set(app)

        Dim test_param As UserParameter

        Dim defaults As New Dictionary(Of UnitsTypeEnum, Object) From _
                {{UnitsTypeEnum.kTextUnits, ""}, _
                 {UnitsTypeEnum.kBooleanUnits, False}, _
                 {UnitsTypeEnum.kUnitlessUnits, 0}}

        'if the parameter doesn't already exist, UserParameters.Item will throw an error
        Try
            test_param = inv_params.Item(n)
        Catch
            Dim default_value = defaults(param_type)

            test_param = inv_params.AddByValue(n, default_value, param_type)
            inv_doc.Update
        End Try
    End Sub

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

    'common method to get the document's custom parameter set
    Public Shared Function get_param_set(ByRef app As Inventor.Application) As UserParameters
        Dim inv_doc As Document = app.ActiveEditDocument
        Dim part_doc As PartDocument
        Dim assm_doc As AssemblyDocument
        Dim inv_params As UserParameters

        'need to treat part and assembly documents slightly differently
        If TypeOf inv_doc Is PartDocument Then
            part_doc = app.ActiveEditDocument
            inv_params = part_doc.ComponentDefinition.Parameters.UserParameters
        ElseIf TypeOf inv_doc Is AssemblyDocument Then
            assm_doc = app.ActiveEditDocument
            inv_params = assm_doc.ComponentDefinition.Parameters.UserParameters
        Else
            'MsgBox("Warning: this is neither a part nor assembly document. Things may misbehave.")
            inv_params = inv_doc.ComponentDefinition.Parameters.UserParameters
        End If

        Return inv_params
    End Function

    'enclose field in quotes, and escape quotes already in the field
    Public Shared Function format_csv_field(ByVal s As String) As String
        Dim s2 As String = Replace(s, """", """""")
        Return """" & s2 & """"
    End Function
End Class
