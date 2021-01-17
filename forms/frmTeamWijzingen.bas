Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10209
    DatasheetFontHeight =11
    ItemSuffix =18
    Left =2520
    Top =1200
    Right =17445
    Bottom =11070
    DatasheetGridlinesColor =15132391
    Filter ="[ToernooiID]=4 and [SessieID] = 15"
    RecSrcDt = Begin
        0xbfce25c5d095e540
    End
    RecordSource ="tblTeamWijzingen"
    Caption ="TeamWijzingen"
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
                    Width =3114
                    Height =969
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Bijschrift16"
                    Caption ="TeamWijzingen"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =3171
                    LayoutCachedHeight =1026
                End
            End
        End
        Begin Section
            Height =6462
            Name ="Details"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =342
                    Height =315
                    ColumnWidth =1701
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="id"
                    ControlSource ="id"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =342
                    LayoutCachedWidth =4593
                    LayoutCachedHeight =657
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =342
                            Top =342
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="id_Bijschrift"
                            Caption ="id"
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
                    ColumnWidth =1080
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"3\";\"2\""
                    Name ="TeamID"
                    ControlSource ="TeamID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblTeams.id, tblTeams.Teamnr, tblTeams.TeamNaam FROM tblTeams; "
                    ColumnWidths ="0;284;2835"
                    GridlineColor =10921638

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
                            Name ="TeamID_Bijschrift"
                            Caption ="TeamID"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =741
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =1071
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =1140
                    Width =1050
                    Height =330
                    ColumnWidth =1050
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Spelnr"
                    ControlSource ="Spelnr"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =1140
                    LayoutCachedWidth =3942
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
                            Name ="Spelnr_Bijschrift"
                            Caption ="Spelnr"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =1140
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =1470
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =1539
                    Width =7260
                    Height =600
                    ColumnWidth =1770
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Speler1"
                    ControlSource ="Speler1"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =1539
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =2139
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =1539
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Speler1_Bijschrift"
                            Caption ="Speler1"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =1539
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =1869
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =2223
                    Width =7260
                    Height =600
                    ColumnWidth =1215
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Speler2"
                    ControlSource ="Speler2"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =2223
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =2823
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =2223
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Speler2_Bijschrift"
                            Caption ="Speler2"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =2223
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =2553
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =2907
                    Width =7260
                    Height =600
                    ColumnWidth =1080
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Speler3"
                    ControlSource ="Speler3"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =2907
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =3507
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =2907
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Speler3_Bijschrift"
                            Caption ="Speler3"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =2907
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =3237
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =3591
                    Width =7260
                    Height =600
                    ColumnWidth =1395
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Speler4"
                    ControlSource ="Speler4"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =3591
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =4191
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =3591
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Speler4_Bijschrift"
                            Caption ="Speler4"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =3591
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =3921
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =4275
                    Width =3465
                    Height =330
                    ColumnWidth =735
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="SessieID"
                    ControlSource ="SessieID"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =4275
                    LayoutCachedWidth =6357
                    LayoutCachedHeight =4605
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =4275
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="SessieID_Bijschrift"
                            Caption ="SessieID"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =4275
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =4605
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3118
                    Top =4932
                    Height =315
                    ColumnWidth =2115
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="ToernooiID"
                    ControlSource ="ToernooiID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblToernooi.ID, tblToernooi.ToernooiNaam FROM tblToernooi ORDER BY tblToe"
                        "rnooi.ToernooiNaam; "
                    ColumnWidths ="0;3969"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =3118
                    LayoutCachedTop =4932
                    LayoutCachedWidth =4819
                    LayoutCachedHeight =5247
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1417
                            Top =4932
                            Width =1110
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift17"
                            Caption ="ToernooiID"
                            GridlineColor =10921638
                            LayoutCachedLeft =1417
                            LayoutCachedTop =4932
                            LayoutCachedWidth =2527
                            LayoutCachedHeight =5247
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
