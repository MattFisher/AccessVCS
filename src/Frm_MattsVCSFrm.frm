Version =19
VersionRequired =19
Checksum =353554403
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ViewsAllowed =1
    TabularFamily =119
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =13209
    DatasheetFontHeight =10
    ItemSuffix =16
    Left =30
    Top =2055
    Right =13815
    Bottom =8385
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa36194295e94e340
    End
    GUID = Begin
        0x7b9efe8439752c4f9cd798cfbe5338b3
    End
    NameMap = Begin
        0x0acc0e5500000000000000000000000000000000000000000c00000003000000 ,
        0x0000000000000000000000000000
    End
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa2050000a1050000a1050000a105000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    PrtDevMode = Begin
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x010400049c009000032f00000100090000000000640001000100c80001000100 ,
        0xc800010000004c65747465720000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x000000000000000000000000000000000000000000000000000000007769646d ,
        0x10000000010000000000000000000000fe0000000100000000000000c8000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x000000000000000000000000
    End
    PrtDevNames = Begin
        0x080036005d000100000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x00000000000000000000000000000000000000000000000000000000004d6963 ,
        0x726f736f667420446f63756d656e7420496d6167696e67205772697465722050 ,
        0x6f72743a00
    End
    OnLoad ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
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
            Height =5669
            BackColor =-2147483633
            Name ="Detail"
            GUID = Begin
                0xb39206f7b1b97b44a388f6bb2b16e9b4
            End
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =170
                    Top =110
                    Width =5553
                    Height =402
                    FontSize =14
                    FontWeight =600
                    Name ="Label0"
                    Caption ="Matt's Access Version Control System"
                    GUID = Begin
                        0xd0c6b765d0eeaa438636a6c10be02c22
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1874
                    Top =1075
                    Width =10548
                    Height =446
                    FontSize =12
                    LeftMargin =57
                    TopMargin =57
                    Name ="C_SourceDirNxt"
                    GUID = Begin
                        0x3b4614f93a58e848a1f61a7e107b1491
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =170
                            Top =1133
                            Width =1560
                            Height =240
                            FontSize =10
                            Name ="Label2"
                            Caption ="Source Directory:"
                            GUID = Begin
                                0xed5971f974e1bd4382d0875162a175b6
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1874
                    Top =1584
                    Width =10548
                    Height =446
                    FontSize =12
                    TabIndex =1
                    LeftMargin =57
                    TopMargin =57
                    Name ="C_RepoUrlNxt"
                    GUID = Begin
                        0x4365dab772841140b455ac7cb63803ce
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =170
                            Top =1644
                            Width =1560
                            Height =240
                            FontSize =10
                            Name ="Label5"
                            Caption ="Repository URL:"
                            GUID = Begin
                                0x70d263f16e56f04da40ab8826aa4b26f
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3628
                    Top =4988
                    Width =1821
                    Height =576
                    TabIndex =2
                    Name ="C_ExportBtn"
                    Caption ="Export Database"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Apply Filter"
                    GUID = Begin
                        0x6e85821928a65043807adfc2800f0989
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1020
                    Top =4988
                    Width =1821
                    Height =576
                    TabIndex =3
                    Name ="C_ImportBtn"
                    Caption ="Import Database"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Apply Filter"
                    GUID = Begin
                        0x4dc3495c1a395b4587c5967291aa6bc7
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2210
                    Top =3414
                    Width =3685
                    Height =340
                    TabIndex =4
                    Name ="C_ContainsOLENxt"
                    GUID = Begin
                        0x04a2d070ec676a40b0896e680233b7a3
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =850
                            Top =3414
                            Width =1335
                            Height =240
                            Name ="Label9"
                            Caption ="Contains OLE?:"
                            GUID = Begin
                                0x33b8fd271c368a409f15582ac0512c00
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2210
                    Top =3874
                    Width =4165
                    Height =1060
                    TabIndex =5
                    Name ="C_EmbeddedPicsNxt"
                    GUID = Begin
                        0x09d16ca50c1a54409768dacc674e519b
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =850
                            Top =3874
                            Width =1335
                            Height =240
                            Name ="Label11"
                            Caption ="Embedded Pics?:"
                            GUID = Begin
                                0x388ed88ae360b5438623de95bcde3787
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6009
                    Top =5045
                    Width =1296
                    Height =501
                    TabIndex =6
                    Name ="C_CheckFormsBtn"
                    Caption ="Check Forms"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Find Next"
                    GUID = Begin
                        0xc70eb2f771e0e84a8eb8bb1e303f0429
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1874
                    Top =566
                    Width =10548
                    Height =446
                    FontSize =12
                    TabIndex =7
                    LeftMargin =57
                    TopMargin =57
                    Name ="C_CurrentDbNxt"
                    GUID = Begin
                        0x1d3da5cc5981324aad0a4f7e89907a95
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =170
                            Top =623
                            Width =1620
                            Height =240
                            FontSize =10
                            Name ="Label14"
                            Caption ="Current Database:"
                            GUID = Begin
                                0x072e5e0b1b31674ab71e72af1716157a
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    PictureType =1
                    Left =12472
                    Top =1077
                    Width =501
                    Height =456
                    FontWeight =700
                    TabIndex =8
                    Name ="C_PickSourceDirBtn"
                    Caption ="..."
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Browse for the Source Directory"
                    GUID = Begin
                        0x3d86d618477ab148b55fed4ed2b41f70
                    End
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub C_CheckFormsBtn_Click()
test_project
End Sub

Private Sub C_ExportBtn_Click()
On Error GoTo ErrProc

SAExportThisDataBase

ExitProc:
Exit Sub
    
ErrProc:
MsgBox Err
End Sub

Private Sub C_ImportBtn_Click()
On Error GoTo ErrProc

SAImportThisDataBase

ExitProc:
Exit Sub
    
ErrProc:
MsgBox Err

End Sub

Private Sub Form_Load()

CheckForOleFields
CheckForEmbeddedImages
Me.C_CurrentDbNxt = CurrentDb.Name()

End Sub

Private Sub CheckForEmbeddedImages()
If FormsHaveEmbeddedImages Then
    Me.C_EmbeddedPicsNxt = "Forms in this database have embedded images." & vbCrLf & _
                           "These will greatly increase the size of the exported files." & vbCrLf & _
                           "It is recommended you include compressed (eg .jpeg) versions of these files" & _
                           "in an ""images"" folder with your database file."
Else
    Me.C_EmbeddedPicsNxt = "There are no embedded pictures in the forms in this database."
End If
End Sub

Private Sub CheckForOleFields()
'If SAContainsOleFields() Then
'    Me.C_ContainsOLENxt = "This database contains OLE fields that must be exported in XML."
'Else
'    Me.C_ContainsOLENxt = "The data in this database may be exported as text or XML as desired."
'End If
End Sub


'Private sub
Private Sub Command12_Click()
On Error GoTo Err_Command12_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command12_Click:
    Exit Sub

Err_Command12_Click:
    MsgBox Err.Description
    Resume Exit_Command12_Click
    
End Sub


Private Sub C_PickSourceDirBtn_Click()
On Error GoTo Err_C_PickSourceDirBtn_Click

'Declare a variable as a FileDialog object.
Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogFolderPicker)
fd.AllowMultiSelect = False

'Declare a variable to contain the path
'of each selected item. Even though the path is a String,
'the variable must be a Variant because For Each...Next
'routines only work with Variants and Objects.
Dim vrtSelectedItem As Variant
'Use the Show method to display the File Picker dialog box and return the user's action.
If fd.Show = -1 Then
    'The user pressed the action button.
    'Get the first string in the FileDialogSelectedItems collection.
    If fd.SelectedItems(1) <> "" Then
        'Perhaps confirm the folder contains interesting stuff here?
        C_SourceDirNxt = fd.SelectedItems(1)
    End If
Else
    'The user pressed Cancel.
End If

Set fd = Nothing

Exit_C_PickSourceDirBtn_Click:
    Exit Sub

Err_C_PickSourceDirBtn_Click:
    MsgBox Err.Description
    Resume Exit_C_PickSourceDirBtn_Click
    
End Sub