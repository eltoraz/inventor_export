' <IsStraightVb>True</IsStraightVb>
Imports Inventor
Imports System.Collections.Generic

Public Module ParameterOps
    'parameters used by multiple modules
    Public shared_params As New Dictionary(Of String, UnitsTypeEnum) From _
            {{"Description", UnitsTypeEnum.kTextUnits}, _
             {"PartNumberToUse", UnitsTypeEnum.kTextUnits}, _
             {"IntermediatePart", UnitsTypeEnum.kBooleanUnits}, _
             {"MaterialsOnly", UnitsTypeEnum.kBooleanUnits}, _
             {"ActiveIsPart", UnitsTypeEnum.kBooleanUnits}, _
             {"FalseParam", UnitsTypeEnum.kBooleanUnits}}

    'master list of parameters created for Epicor module
    Public epicor_params As New Dictionary(Of String, UnitsTypeEnum) From _
            {{"ProdCode", UnitsTypeEnum.kTextUnits}, _
             {"ClassID", UnitsTypeEnum.kTextUnits}, _
             {"UsePartRev", UnitsTypeEnum.kBooleanUnits}, _
             {"MfgComment", UnitsTypeEnum.kTextUnits}, _
             {"PurComment", UnitsTypeEnum.kTextUnits}, _
             {"TrackSerialNum", UnitsTypeEnum.kBooleanUnits}, _
             {"RevDescription", UnitsTypeEnum.kTextUnits}}

    'valid species parts will use (encoded in parameter names)
    Public species_list = New String() {"Ash", "Birch-Baltic", "Cherry", _
                                 "Maple-Hard", "Maple-Soft", "Oak-Red", "Oak-White", _
                                 "Pine", "Poplar", "Walnut", "Hardware", "Birch-White"}

    'master list of parameters created for Quoting module
    'empty ArrayList represents user-entered field
    Public quoting_params As New Dictionary(Of String, Tuple(Of UnitsTypeEnum, ArrayList)) From _
            {{"QuotingSpreadsheet", Tuple.Create(UnitsTypeEnum.kTextUnits, _
                        New ArrayList())}, _
             {"FinishedThickness", Tuple.Create(UnitsTypeEnum.kUnitlessUnits, _
                        New ArrayList())}, _
             {"Width", Tuple.Create(UnitsTypeEnum.kUnitlessUnits, _
                        New ArrayList())}, _
             {"Length", Tuple.Create(UnitsTypeEnum.kUnitlessUnits, _
                        New ArrayList())}, _
             {"WidthSpec", Tuple.Create(UnitsTypeEnum.kTextUnits, _
                        New ArrayList() From {"RTW", "PWT"})}, _
             {"LengthSpec", Tuple.Create(UnitsTypeEnum.kTextUnits, _
                        New ArrayList() From {"RTL", "PET"})}, _
             {"QtyPerUnit", Tuple.Create(UnitsTypeEnum.kUnitlessUnits, _
                        New ArrayList())}, _
             {"NestedQty", Tuple.Create(UnitsTypeEnum.kUnitlessUnits, _
                        New ArrayList())}, _
             {"SandingSpec", Tuple.Create(UnitsTypeEnum.kTextUnits, _
                        New ArrayList() From {"S2S_150", "S4S_150", "S0S", _
                                              "SDIA", "SSF", "NSF"})}, _
             {"GrainDirection", Tuple.Create(UnitsTypeEnum.kTextUnits, _
                        New ArrayList() From {"GDL", "GDN"})}, _
             {"CertifiedClass", Tuple.Create(UnitsTypeEnum.kTextUnits, _
                        New ArrayList() From {"FSC", "FSC_CARB_93120", "NCA"})}, _
             {"WoodSpecies", Tuple.Create(UnitsTypeEnum.kTextUnits, _
                        New ArrayList())}, _
             {"GlueUpSpec", Tuple.Create(UnitsTypeEnum.kTextUnits, _
                        New ArrayList() From {"SOLID", "SOLID_GLUE", "PANEL", _
                                              "FACE_GLUE", "PLYWOOD", "DOWEL"})}, _
             {"ColorSpec", Tuple.Create(UnitsTypeEnum.kTextUnits, _
                        New ArrayList() From {"USEL", "WM"})}, _
             {"GradeSpec", Tuple.Create(UnitsTypeEnum.kTextUnits, _
                        New ArrayList() From {"C1F", "C2F", "USEL"})}, _
             {"CustomSpec", Tuple.Create(UnitsTypeEnum.kTextUnits, _
                        New ArrayList() From {"E2E", "E4E", "C2E", "C4E", _
                                              "BN1E", "BN2E", "Bend", "Custom"})}, _
             {"CustomDetails", Tuple.Create(UnitsTypeEnum.kTextUnits, _
                        New ArrayList())}}

    'species-specific option lists for color spec in Quoting module
    Public quoting_color_specs As New Dictionary(Of String, ArrayList) From _
            {{"Cherry", New ArrayList() From {"R1F", "R2F"}}, _
             {"Maple-Hard", New ArrayList() From {"W1F", "W2F"}}, _
             {"Maple-Soft", New ArrayList() From {"W1F", "W2F"}}, _
             {"Walnut", New ArrayList() From {"B1F", "B2F"}}}


    '----------------------Methods-------------------------------------------

    'initialize parameter `n` as type `param_type`
    Public Sub create_param(ByVal n As String, ByVal param_type As UnitsTypeEnum, _
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

    'create all parameters needed for the suite
    'note: this doesn't necessarily initialize them to default values that make sense!
    Public Sub create_all_params(ByRef inv_app As Inventor.Application)
        'shared
        For Each kvp As KeyValuePair(Of String, UnitsTypeEnum) in shared_params
            create_param(kvp.Key, kvp.Value, inv_app)
        Next

        'Epicor module
        For Each kvp As KeyValuePair(Of String, UnitsTypeEnum) in epicor_params
            create_param(kvp.Key, kvp.Value, inv_app)
        Next

        'species module
        For Each s As String in species_list
            'note: Inventor parameters don't support spaces or special characters, so
            'need to do a character substitution on the `-`, then switch back when
            'converting to iproperties
            Dim subst As String = Replace(s, "-", "4")
            create_param("Flag" & subst, UnitsTypeEnum.kBooleanUnits, inv_app)
            create_param("Part" & subst, UnitsTypeEnum.kTextUnits, inv_app)
            create_param("ExportedPart" & subst, UnitsTypeEnum.kBooleanUnits, inv_app)

            '"Hardware" doesn't have a material associated
            If Not String.Equals(s, "Hardware") Then
                create_param("FlagMat" & subst, UnitsTypeEnum.kBooleanUnits, inv_app)
                create_param("Mat" & subst, UnitsTypeEnum.kTextUnits, inv_app)
                create_param("ExportedMat" & subst, UnitsTypeEnum.kBooleanUnits, inv_app)
            End If
        Next

        'quoting module
        For Each kvp As KeyValuePair(Of String, Tuple(Of UnitsTypeEnum, ArrayList)) In quoting_params
            create_param(kvp.Key, kvp.Value.Item1, inv_app)
        Next
        'create color spec parameters for each species
        For Each s As String in species_list
            Dim subst As String = Replace(s, "-", "4")
            create_param("ColorSpec" & subst, UnitsTypeEnum.kTextUnits, inv_app)
        Next

        'set some parameters to safe defaults (eg, prevent div by 0)
        ' (but only if they haven't already been modified in their respective modules!)
        Dim inv_params = get_param_set(inv_app)
        If inv_params.Item("NestedQty").Value = 0 Then inv_params.Item("NestedQty").Value = 1
    End Sub

    'common method to get the document's custom parameter set
    Public Function get_param_set(ByRef app As Inventor.Application) As UserParameters
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
End Module