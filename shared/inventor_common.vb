' <IsStraightVb>True</IsStraightVb>
'Inventor parameter/iProperty manipulation functions
Imports Inventor

Public Class InventorOps
    'initialize parameter `n` as type `paramType`
    Public Shared Sub create_param(ByVal n As String, ByVal paramType As UnitsTypeEnum, _
                                   ByRef app As Inventor.Application)
        dim invDoc As Document = app.ActiveDocument

        Dim invParams As UserParameters = invDoc.Parameters.UserParameters

        Dim TestParam As UserParameter

        'if the parameter doesn't already exist, UserParameters.Item will throw an error
        Try
            TestParam = invParams.Item(n)
        Catch
            Dim defaultValue
            If paramType = UnitsTypeEnum.kTextUnits Then
                defaultValue = ""
            ElseIf paramType = UnitsTypeEnum.kBooleanUnits Then
                defaultValue = False
            ElseIf paramType = UnitsTypeEnum.kUnitlessUnits Then
                defaultValue = 0
            End If

            TestParam = invParams.AddByValue(n, defaultValue, paramType)
            invDoc.Update
        End Try
    End Sub

    'update iProperty `n` with value `paramVal`, creating it if it doesn't exist
    Public Shared Sub update_prop(ByVal n As String, ByVal paramVal As Object, _
                                  ByRef app As Inventor.Application)
        'get the custom property collection
        Dim invDoc As Document = app.ActiveDocument
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
End Class
