' <IsStraightVb>True</IsStraightVb>
Imports System.Text.RegularExpressions
Imports Inventor
Imports Autodesk.iLogic.Interfaces

Public Class SpeciesOps
    'regular expressions to match the part number format WP-ZZ-123, and
    ' material part number MX-ZZ-123
    Public Shared part_pattern As String = "[Ww][Pp]-[a-zA-Z]{2}-[0-9]{3}"
    Public Shared mat_pattern As String = "[Mm][lhftbpLHFTBP]-[a-zA-Z]{2}-[0-9]{3}"

    Public Shared Function select_active_part(ByRef app As Inventor.Application, _
                                              ByRef inv_params As UserParameters, _
                                              ByRef species_list() As String, _
                                              ByRef form_obj As IiLogicForm, _
                                              ByRef vb_obj As ILowLevelSupport, _
                                              ByRef multivalue_obj As IMultiValueParam, _
                                              ByVal materials_only As Boolean) _
                                              As FormResult
        'select the part we'll be working with here
        Dim active_parts As New ArrayList()
        Dim no_species As Boolean = False

        Dim form_result As FormResult = FormResult.OK

        Do
            Try
                For Each s As String in species_list
                    Dim subst As String = Replace(s, "-", "4")

                    Dim flag_param As Parameter = inv_params.Item("Flag" & subst)
                    Dim flag_value = flag_param.Value

                    If flag_value Then
                        'add active parts and materials to the list to present to the user
                        If Not materials_only Then
                            Dim part_param As Parameter = inv_params.Item("Part" & subst)
                            Dim part_value As String = part_param.Value
                            Dim part_entry As String = part_value & " - " & s
                            active_parts.Add(part_entry)
                        End If

                        If StrComp(s, "Hardware") <> 0 AndAlso inv_params.Item("FlagMat" & subst).Value Then
                            Dim mat_param As Parameter = inv_params.Item("Mat" & subst)
                            Dim mat_value As String = mat_param.Value
                            Dim mat_entry As String = mat_value & " - " & s
                            active_parts.Add(mat_entry)
                        End If
                    End If
                Next
            Catch e As ArgumentException
                'exception thrown when using UserParameters.Item above and trying to
                'get a parameter that doesn't exist
                no_species = True
            End Try
            
            'check whether the species parameters have been created but none selected
            If active_parts.Count = 0 Then
                no_species = True
            Else
                no_species = False
            End If

            'can't proceed if there isn't a part number for at least one species
            If no_species Then
                form_result = form_obj.ShowGlobal("epicor_13launch_species", FormMode.Modal).Result

                If form_result = FormResult.None Then
                    Return form_result
                End If
            End If
        Loop While no_species

        multivalue_obj.List("PartNumberToUse") = active_parts

        Dim part_selected As Boolean = False
        Dim pn As String = ""
        Do
            form_result = form_obj.ShowGlobal("epicor_15part_select", FormMode.Modal).Result

            If form_result = FormResult.Cancel OrElse form_result = FormResult.None Then
                Return form_result
            End If

            pn = inv_params.Item("PartNumberToUse").Value
            If StrComp(pn, "") <> 0 Then
                part_selected = True
            Else
                MsgBox("Please select a part to continue.")
                vb_obj.RunExternalRule("dummy.vb")
            End If
        Loop While Not part_selected

        'set some parameters based on the part selected
        Dim part_fields As Tuple(Of String, String, String) = unpack_pn(pn)
        inv_params.Item("PartType").Value = part_fields.Item2
        If String.Equals(part_fields.Item2, "P") Then
            inv_params.Item("ActiveIsPart").Value = False
        Else
            inv_params.Item("ActiveIsPart").Value = True
        End If

        Return form_result
    End Function

    'parse the part number into a tuple in the format (partnumber, parttype, species)
    '(or a tuple of empty strings if the input doesn't match the expected format)
    Public Shared Function unpack_pn(ByVal pn As String) As Tuple(Of String, String, String)
        'pn is in the format `MX-ZZ-123 - Species` for raw materials
        'manufactured parts will usually be `WP-ZZ-123`, but that only applies
        ' for intermediate parts (parts that are specified by customers will
        ' have customer-specified pn that may not fit this format)

        'use regex match groups to capture the part number and species
        'infer the part type from which pattern matches
        Dim general_part_patter As String = "[\w\-]+"
        Dim part_grouped As String = "^(" & general_part_pattern & ") - (\w+-?\w+)$"
        Dim part_regex As New Regex(part_grouped)
        Dim mat_grouped As String = "^(" & mat_pattern & ") - (\w+-?\w+)$"
        Dim mat_regex As New Regex(mat_grouped)

        Dim part_num, part_type, part_species As String

        Dim p_match As Match = part_regex.Match(pn)
        Dim m_match As Match = mat_regex.Match(pn)
        If p_match.Success Then
            part_num = p_match.Groups(1).Value
            part_type = "M"
            part_species = p_match.Groups(2).Value
        ElseIf m_match.Success Then
            part_num = m_match.Groups(1).Value
            part_type = "P"
            part_species = m_match.Groups(2).Value
        Else
            part_num = ""
            part_type = ""
            part_species = ""
        End If

        Return Tuple.Create(part_num, part_type, part_species)

    End Function
End Class
