Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10209
    DatasheetFontHeight =11
    ItemSuffix =21
    Right =12300
    Bottom =10245
    DatasheetGridlinesColor =15132391
    Filter ="[ToernooiID]=1 and [SessieID] = 10"
    RecSrcDt = Begin
        0x34702d88f391e540
    End
    RecordSource ="tblUitslagen"
    Caption ="frmTeamUitslagen"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =1026
            BackColor =15064278
            Name ="Formulierkoptekst"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =57
                    Top =57
                    Width =3726
                    Height =969
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Bijschrift20"
                    Caption ="frmTeamUitslagen"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =3783
                    LayoutCachedHeight =1026
                End
            End
        End
        Begin Section
            Height =4647
            Name ="Details"
            AutoHeight =1
            AlternateBackColor =15658734
            BackThemeColorIndex =1
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2892
                    Top =342
                    Width =3660
                    Height =330
                    ColumnWidth =975
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"3\";\"2\""
                    Name ="SessieID"
                    ControlSource ="SessieID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblSessie.id, tblSessie.Sessienr FROM tblSessie; "
                    ColumnWidths ="0;1134"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =2892
                    LayoutCachedTop =342
                    LayoutCachedWidth =6552
                    LayoutCachedHeight =672
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =342
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="SessieID_Bijschrift"
                            Caption ="Sessie"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =342
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =672
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =2892
                    Top =741
                    Width =3660
                    Height =330
                    ColumnWidth =1140
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"3\";\"2\""
                    Name ="TeamIDThuis"
                    ControlSource ="TeamIDThuis"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblTeams.id, tblTeams.Teamnr, tblTeams.TeamNaam FROM tblTeams; "
                    ColumnWidths ="0;284;1701"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =2892
                    LayoutCachedTop =741
                    LayoutCachedWidth =6552
                    LayoutCachedHeight =1071
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =741
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="TeamIDThuis_Bijschrift"
                            Caption ="Thuis"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =741
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =1071
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =2892
                    Top =1140
                    Width =3660
                    Height =330
                    ColumnWidth =1110
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"3\";\"2\""
                    Name ="TeamIDUit"
                    ControlSource ="TeamIDUit"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblTeams.id, tblTeams.Teamnr, tblTeams.TeamNaam FROM tblTeams; "
                    ColumnWidths ="0;284;1701"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =2892
                    LayoutCachedTop =1140
                    LayoutCachedWidth =6552
                    LayoutCachedHeight =1470
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =1140
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="TeamIDUit_Bijschrift"
                            Caption ="Uit"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =1140
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =1470
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =1539
                    Width =1530
                    Height =330
                    ColumnWidth =690
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Wedstrijdnr"
                    ControlSource ="Wedstrijdnr"
                    StatusBarText ="Bij meerdere wedstrijd per sessie rangnummer 1 .. Aantalwedstrijden per sessie"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =1539
                    LayoutCachedWidth =4422
                    LayoutCachedHeight =1869
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =1539
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Wedstrijdnr_Bijschrift"
                            Caption ="Wedstrijdnr"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =1539
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =1869
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =1938
                    Width =1050
                    Height =330
                    ColumnWidth =1185
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ImpsThuis"
                    ControlSource ="ImpsThuis"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =1938
                    LayoutCachedWidth =3942
                    LayoutCachedHeight =2268
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =1938
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ImpsThuis_Bijschrift"
                            Caption ="ImpsThuis"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =1938
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =2268
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =2337
                    Width =1050
                    Height =330
                    ColumnWidth =1050
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ImpsUit"
                    ControlSource ="ImpsUit"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =2337
                    LayoutCachedWidth =3942
                    LayoutCachedHeight =2667
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =2337
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ImpsUit_Bijschrift"
                            Caption ="ImpsUit"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =2337
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =2667
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =2736
                    Width =3660
                    Height =330
                    ColumnWidth =1380
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="VpsThuis"
                    ControlSource ="VpsThuis"
                    Format ="Standard"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =2736
                    LayoutCachedWidth =6552
                    LayoutCachedHeight =3066
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =2736
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="VpsThuis_Bijschrift"
                            Caption ="VpsThuis"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =2736
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =3066
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =3135
                    Width =3660
                    Height =330
                    ColumnWidth =1305
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="VpsUit"
                    ControlSource ="VpsUit"
                    Format ="Standard"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =3135
                    LayoutCachedWidth =6552
                    LayoutCachedHeight =3465
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =3135
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="VpsUit_Bijschrift"
                            Caption ="VpsUit"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =3135
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =3465
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2892
                    Top =3534
                    Width =3660
                    Height =330
                    ColumnWidth =3600
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="ToernooiID"
                    ControlSource ="ToernooiID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblToernooi.ID, tblToernooi.ToernooiNaam FROM tblToernooi; "
                    ColumnWidths ="0;1701"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =2892
                    LayoutCachedTop =3534
                    LayoutCachedWidth =6552
                    LayoutCachedHeight =3864
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =3534
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ToernooiID_Bijschrift"
                            Caption ="Toernooi"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =3534
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =3864
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =3933
                    Width =7260
                    Height =600
                    ColumnWidth =1185
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Tafel"
                    ControlSource ="Tafel"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =3933
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =4533
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =3933
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Tafel_Bijschrift"
                            Caption ="Tafel"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =3933
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =4263
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="Formuliervoettekst"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(Cancel As Integer)
    If CurrentProject.AllForms("Start_VT").IsLoaded = True Then
        Me.Filter = "[ToernooiID]=" & lngToernooi & " and [SessieID] = " & lngSessie
        Me.FilterOn = True
    End If
End Sub
