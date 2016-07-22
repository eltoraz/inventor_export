' <IsStraightVb>True</IsStraightVb>
Imports Inventor

Public Class QuotingOps
    'master list of parameters created for Quoting module
    Public Shared param_list As New Dictionary(Of String, Tuple(Of UnitsTypeEnum, ArrayList)) From _
            {{"PartDescription", Tuple.Create(UnitsTypeEnum.kTextUnits, _
                        New ArrayList())}, _
             {"FinishedThickness", Tuple.Create(UnitsTypeEnum.kInchLengthUnits, _
                        New ArrayList() From {0.75, 1.00, 1.25, 1.75, 1.75, 2.25, 2.75})}, _
             {"Width", Tuple.Create(UnitsTypeEnum.kInchLengthUnits, _
                        New ArrayList())}, _
             {"Length", Tuple.Create(UnitsTypeEnum.kInchLengthUnits, _
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
                        New ArrayList() From {"USEL", "WM", "R1F", "R2F", "W1F", _
                                              "W2F", "B1F", "B2F", "I", "U"})}, _
             {"GradeSpec", Tuple.Create(UnitsTypeEnum.kTextUnits, _
                        New ArrayList() From {"C1F", "C2F", "USEL"})}}
End Class
