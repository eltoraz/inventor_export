' <IsStraightVb>True</IsStraightVb>
Imports Inventor

Public Class ParameterLists
    'parameters used by multiple modules
    Public Shared shared_params As New Dictionary(Of String, UnitsTypeEnum) From _
            {{"PartType", UnitsTypeEnum.kTextUnits}, _
             {"PartNumberToUse", UnitsTypeEnum.kTextUnits}, _
             {"IntermediatePart", UnitsTypeEnum.kBooleanUnits}, _
             {"MaterialsOnly", UnitsTypeEnum.kBooleanUnits}, _
             {"ActiveIsPart", UnitsTypeEnum.kBooleanUnits}, _
             {"FalseParam", UnitsTypeEnum.kBooleanUnits}}

    'master list of parameters created for Epicor module
    Public Shared epicor_params As New Dictionary(Of String, UnitsTypeEnum) From _
            {{"Description", UnitsTypeEnum.kTextUnits}, _
             {"ProdCode", UnitsTypeEnum.kTextUnits}, _
             {"ClassID", UnitsTypeEnum.kTextUnits}, _
             {"UsePartRev", UnitsTypeEnum.kBooleanUnits}, _
             {"MfgComment", UnitsTypeEnum.kTextUnits}, _
             {"PurComment", UnitsTypeEnum.kTextUnits}, _
             {"TrackSerialNum", UnitsTypeEnum.kBooleanUnits}, _
             {"RevDescription", UnitsTypeEnum.kTextUnits}}

    'master list of parameters created for Quoting module
    'empty ArrayList represents user-entered field
    Public Shared quoting_params As New Dictionary(Of String, Tuple(Of UnitsTypeEnum, ArrayList)) From _
            {{"QuotingSpreadsheet", Tuple.Create(UnitsTypeEnum.kTextUnits, _
                        New ArrayList())}, _
             {"FinishedThickness", Tuple.Create(UnitsTypeEnum.kUnitlessUnits, _
                        New ArrayList() From {0.75, 1.00, 1.25, 1.75, 1.75, 2.25, 2.75})}, _
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

    Public Shared quoting_color_specs As New Dictionary(Of String, ArrayList) From _
            {{"Cherry", New ArrayList() From {"R1F", "R2F"}}, _
             {"Maple-Hard", New ArrayList() From {"W1F", "W2F"}}, _
             {"Maple-Soft", New ArrayList() From {"W1F", "W2F"}}, _
             {"Walnut", New ArrayList() From {"B1F", "B2F"}}}
End Class
