Version =19
VersionRequired =19
Checksum =1168207333
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7021
    DatasheetFontHeight =10
    ItemSuffix =10
    Left =3900
    Top =2730
    Right =11205
    Bottom =5670
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x4c560af713a3e340
    End
    GUID = Begin
        0x85154994d9fab1409e67bb8c2fcf2681
    End
    NameMap = Begin
        0x0acc0e5500000000807165f181350144a66b0925849c1777000000000ed21703 ,
        0x14a3e34000008e00010000005400610073006b007300000000000000636ce3b1 ,
        0xa451454db7efd0e2cce420b607000000807165f181350144a66b0925849c1777 ,
        0x5400610073006b0049004400000000000000af1342a7ea2ac84ba4c7496dbfae ,
        0xb8d807000000807165f181350144a66b0925849c17775400610073006b004400 ,
        0x650073006300720069007000740069006f006e00000000000000b612a2c55ef3 ,
        0xc646be7893859b9a776107000000807165f181350144a66b0925849c17775300 ,
        0x7400610072007400440061007400650000000000000024f9c43bd8f5c949b558 ,
        0x634f4d18f97907000000807165f181350144a66b0925849c177745006e006400 ,
        0x4400610074006500000000000000de9c14b25c027c48ab9bbcccdaf9d9170700 ,
        0x0000807165f181350144a66b0925849c17774e006f0074006500730000000000 ,
        0x0000000000000000000000000000000000000c00000003000000000000000000 ,
        0x0000000000000000
    End
    RecordSource ="Tasks"
    Caption ="Tasks"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =2955
            BackColor =-2147483633
            Name ="Detail"
            GUID = Begin
                0xfd33932152dd53478a27cb0c9bcfc8db
            End
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    Name ="TaskID"
                    ControlSource ="TaskID"
                    GUID = Begin
                        0xbcedea09026ce44e87d410410013bc6f
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =615
                            Height =240
                            Name ="Label1"
                            Caption ="Task ID"
                            GUID = Begin
                                0x979fa50e509c8b48b53b3b45edac84ab
                            End
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    Width =4950
                    Height =795
                    TabIndex =1
                    Name ="TaskDescription"
                    ControlSource ="TaskDescription"
                    GUID = Begin
                        0xbfff0a8a0155ca408298d20e98f05aa7
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =1245
                            Height =240
                            Name ="Label3"
                            Caption ="Task Description"
                            GUID = Begin
                                0x135aaf2e8916874bb577155aea7fdce2
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1927
                    Top =1360
                    TabIndex =2
                    Name ="StartDate"
                    ControlSource ="StartDate"
                    Format ="Short Date"
                    InputMask ="99/99/00;0"
                    GUID = Begin
                        0x6d864cb149d2e141948822fcb5185626
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1360
                            Width =825
                            Height =240
                            Name ="Label5"
                            Caption ="Start Date"
                            GUID = Begin
                                0x4d48b1da5569e14884ca7a0a6f5f488d
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1927
                    Top =1700
                    TabIndex =3
                    Name ="EndDate"
                    ControlSource ="EndDate"
                    Format ="Short Date"
                    InputMask ="99/99/00;0"
                    GUID = Begin
                        0x83019e482a935043a27e13bd8dc44414
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1700
                            Width =735
                            Height =240
                            Name ="Label7"
                            Caption ="End Date"
                            GUID = Begin
                                0x79ab59918605614f8239f1258d396100
                            End
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1927
                    Top =2040
                    Width =4950
                    Height =795
                    TabIndex =4
                    Name ="Notes"
                    ControlSource ="Notes"
                    GUID = Begin
                        0x21c44c0c8d8d7c45960a6056bd98d807
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2040
                            Width =495
                            Height =240
                            Name ="Label9"
                            Caption ="Notes"
                            GUID = Begin
                                0x3acf6319aa73d346bc6b394ffa9ad1a0
                            End
                        End
                    End
                End
            End
        End
    End
End