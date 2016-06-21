' <IsStraightVb>True</IsStraightVb>

Public Class Species
    'master list of species
    'NOTE: "Hardware" is a special case for our purposes - it's the only one that
    '      won't have a "Material" associated with it
    Public Shared species_list = New String() {"Ash", "Birch-Baltic", "Cherry", _
                                 "Maple-Hard", "Maple-Soft", "Oak-Red", "Oak-White", _
                                 "Pine", "Poplar", "Walnut", "Hardware"}
End Class
