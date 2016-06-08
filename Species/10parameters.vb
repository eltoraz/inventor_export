AddVbFile "inventor_common.vb"      'InventorOps.create_param

'create 3 parameters per supported species: one to flag whether it's in use,
'1 for the part, and 1 for the raw material
Sub Main()
    Dim species As New ArrayList()

    'note: Inventor parameters don't support spaces or special characters, so
    'need to do a character substitution on the `-`, then switch back when
    'converting to iproperties
    species.Add("Ash")
    species.Add("Birch-Baltic")
    species.Add("Cherry")
    species.Add("Maple-Hard")
    species.Add("Maple-Soft")
    species.Add("Oak-Red")
    species.Add("Oak-White")
    species.Add("Pine")
    species.Add("Poplar")
    species.Add("Walnut")

    For Each s As String in species:
        Dim subst As String = Replace(s, "-", "4")
        InventorOps.create_param("Flag" & subst, UnitsTypeEnum.kBooleanUnits)
        InventorOps.create_param("Part" & subst, UnitsTypeEnum.kTextUnits)
        InventorOps.create_param("Mat" & subst, UnitsTypeEnum.kTextUnits)
    Next

    'special case: "Hardware", to be handled on an individual basis
    Dim hw As String = "Hardware"
    InventorOps.create_param("Flag" & hw, UnitsTypeEnum.kBooleanUnits)
    InventorOps.create_param("Part" & hw, UnitsTypeEnum.kTextUnits)
End Sub
