AddVbFile "inventor_common.vb"      'InventorOps.create_param
AddVbFile "species_list.vb"         'Species.species_list

'create 3 parameters per supported species: one to flag whether it's in use,
'1 for the part, and 1 for the raw material
Sub Main()
    Dim inv_app As Inventor.Application = ThisApplication
    For Each s As String in Species.species_list
        'note: Inventor parameters don't support spaces or special characters, so
        'need to do a character substitution on the `-`, then switch back when
        'converting to iproperties
        Dim subst As String = Replace(s, "-", "4")
        InventorOps.create_param("Flag" & subst, UnitsTypeEnum.kBooleanUnits, inv_app)
        InventorOps.create_param("Part" & subst, UnitsTypeEnum.kTextUnits, inv_app)

        '"Hardware" doesn't have a material associated
        If StrComp(s, "Hardware") <> 0 Then
            InventorOps.create_param("Mat" & subst, UnitsTypeEnum.kTextUnits, inv_app)
        End If
    Next
End Sub
