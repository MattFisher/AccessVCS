Version =19
VersionRequired =19
Checksum =-28133460
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
        0x700effd05b92e340
    End
    GUID = Begin
        0x85154994d9fab1409e67bb8c2fcf2681
    End
    NameMap = Begin
        0x0acc0e55000000000d39d6f66676c2468601ed3ca30986d000000000c55b36cc ,
        0x5b92e34000008e00010206005400610073006b0073000000000000007f25fab3 ,
        0x7a2b40418e3cb3513070f05d070000000d39d6f66676c2468601ed3ca30986d0 ,
        0x5400610073006b0049004400000000000000555e9f5eff22b847a68c868c0bec ,
        0x05b9070000000d39d6f66676c2468601ed3ca30986d05400610073006b004400 ,
        0x650073006300720069007000740069006f006e00000000000000f750dc9395ea ,
        0xf04ebdbe2def12d57717070000000d39d6f66676c2468601ed3ca30986d05300 ,
        0x74006100720074004400610074006500000000000000864273188f532442b946 ,
        0x9796d963dc69070000000d39d6f66676c2468601ed3ca30986d045006e006400 ,
        0x44006100740065000000000000008176c3ae0ebc8f48a14513d240a41ecd0700 ,
        0x00000d39d6f66676c2468601ed3ca30986d04e006f0074006500730000000000 ,
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