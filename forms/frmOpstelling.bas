Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10209
    DatasheetFontHeight =11
    ItemSuffix =36
    Left =2520
    Top =1200
    Right =17445
    Bottom =11070
    DatasheetGridlinesColor =15132391
    Filter ="[ToernooiID]=5 and [SessieID] = 16"
    OrderBy ="[tblOpstelling].[ToernooiID], [tblOpstelling].[Sessie], [tblOpstelling].[Teamnr]"
    RecSrcDt = Begin
        0x7168df213190e540
    End
    RecordSource ="tblOpstelling"
    Caption ="frmOpstelling"
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
                    Width =2772
                    Height =969
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Bijschrift32"
                    Caption ="frmOpstelling"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =2829
                    LayoutCachedHeight =1026
                End
            End
        End
        Begin Section
            Height =7911
            Name ="Details"
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =342
                    Height =315
                    ColumnWidth =1701
                    ColumnOrder =0
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
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =741
                    Width =1530
                    Height =330
                    ColumnWidth =975
                    ColumnOrder =1
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ToernooiID"
                    ControlSource ="ToernooiID"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =741
                    LayoutCachedWidth =4422
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
                            Name ="ToernooiID_Bijschrift"
                            Caption ="ToernooiID"
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
                    Width =960
                    Height =330
                    ColumnWidth =1005
                    ColumnOrder =2
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Sessie"
                    ControlSource ="Sessie"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =1140
                    LayoutCachedWidth =3852
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
                            Name ="Sessie_Bijschrift"
                            Caption ="Sessie"
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
                    Width =1050
                    Height =330
                    ColumnWidth =2295
                    ColumnOrder =3
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Teamnr"
                    ControlSource ="Teamnr"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =1539
                    LayoutCachedWidth =3942
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
                            Name ="Teamnr_Bijschrift"
                            Caption ="Teamnr"
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
                    Top =1938
                    Width =7260
                    Height =600
                    ColumnWidth =1365
                    ColumnOrder =5
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Speler1"
                    ControlSource ="Speler1"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =1938
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =2538
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =1938
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Speler1_Bijschrift"
                            Caption ="Speler1"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =1938
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =2268
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =2622
                    Width =7260
                    Height =600
                    ColumnWidth =1440
                    ColumnOrder =6
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Speler2"
                    ControlSource ="Speler2"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =2622
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =3222
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =2622
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Speler2_Bijschrift"
                            Caption ="Speler2"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =2622
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =2952
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =3306
                    Width =7260
                    Height =600
                    ColumnWidth =1230
                    ColumnOrder =7
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Speler3"
                    ControlSource ="Speler3"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =3306
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =3906
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =3306
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Speler3_Bijschrift"
                            Caption ="Speler3"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =3306
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =3636
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =3990
                    Width =7260
                    Height =600
                    ColumnWidth =1500
                    ColumnOrder =8
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Speler4"
                    ControlSource ="Speler4"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =3990
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =4590
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =3990
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Speler4_Bijschrift"
                            Caption ="Speler4"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =3990
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =4320
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =4674
                    Width =1050
                    Height =330
                    ColumnWidth =765
                    ColumnOrder =9
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Wedstrijd1"
                    ControlSource ="Wedstrijd1"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =4674
                    LayoutCachedWidth =3942
                    LayoutCachedHeight =5004
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =4674
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Wedstrijd1_Bijschrift"
                            Caption ="Wedstrijd1"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =4674
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =5004
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =5073
                    Width =1050
                    Height =330
                    ColumnWidth =765
                    ColumnOrder =10
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Wedstrijd2"
                    ControlSource ="Wedstrijd2"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =5073
                    LayoutCachedWidth =3942
                    LayoutCachedHeight =5403
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =5073
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Wedstrijd2_Bijschrift"
                            Caption ="Wedstrijd2"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =5073
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =5403
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =5472
                    Width =1050
                    Height =330
                    ColumnWidth =645
                    ColumnOrder =11
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Wedstrijd3"
                    ControlSource ="Wedstrijd3"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =5472
                    LayoutCachedWidth =3942
                    LayoutCachedHeight =5802
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =5472
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Wedstrijd3_Bijschrift"
                            Caption ="Wedstrijd3"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =5472
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =5802
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =5871
                    Width =1050
                    Height =330
                    ColumnWidth =465
                    ColumnOrder =12
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Wedstrijd4"
                    ControlSource ="Wedstrijd4"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =5871
                    LayoutCachedWidth =3942
                    LayoutCachedHeight =6201
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =5871
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Wedstrijd4_Bijschrift"
                            Caption ="Wedstrijd4"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =5871
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =6201
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =6270
                    Width =1050
                    Height =330
                    ColumnWidth =450
                    ColumnOrder =13
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Wedstrijd5"
                    ControlSource ="Wedstrijd5"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =6270
                    LayoutCachedWidth =3942
                    LayoutCachedHeight =6600
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =6270
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Wedstrijd5_Bijschrift"
                            Caption ="Wedstrijd5"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =6270
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =6600
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =6669
                    Width =1050
                    Height =330
                    ColumnWidth =390
                    ColumnOrder =14
                    TabIndex =13
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Wedstrijd6"
                    ControlSource ="Wedstrijd6"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =6669
                    LayoutCachedWidth =3942
                    LayoutCachedHeight =6999
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =6669
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Wedstrijd6_Bijschrift"
                            Caption ="Wedstrijd6"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =6669
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =6999
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =7068
                    Width =1050
                    Height =330
                    ColumnWidth =450
                    ColumnOrder =15
                    TabIndex =14
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Wedstrijd7"
                    ControlSource ="Wedstrijd7"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =7068
                    LayoutCachedWidth =3942
                    LayoutCachedHeight =7398
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =7068
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Wedstrijd7_Bijschrift"
                            Caption ="Wedstrijd7"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =7068
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =7398
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =7467
                    Width =1050
                    Height =330
                    ColumnWidth =360
                    ColumnOrder =16
                    TabIndex =15
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Wedstrijd8"
                    ControlSource ="Wedstrijd8"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =7467
                    LayoutCachedWidth =3942
                    LayoutCachedHeight =7797
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =7467
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Wedstrijd8_Bijschrift"
                            Caption ="Wedstrijd8"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =7467
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =7797
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7653
                    Top =5102
                    Height =315
                    ColumnWidth =285
                    ColumnOrder =17
                    TabIndex =16
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Wedstrijd9"
                    ControlSource ="Wedstrijd9"
                    GridlineColor =10921638

                    LayoutCachedLeft =7653
                    LayoutCachedTop =5102
                    LayoutCachedWidth =9354
                    LayoutCachedHeight =5417
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5952
                            Top =5102
                            Width =1095
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift33"
                            Caption ="Wedstrijd9"
                            GridlineColor =10921638
                            LayoutCachedLeft =5952
                            LayoutCachedTop =5102
                            LayoutCachedWidth =7047
                            LayoutCachedHeight =5417
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7143
                    Top =5725
                    Height =315
                    ColumnWidth =480
                    ColumnOrder =18
                    TabIndex =17
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Wedstrijd10"
                    ControlSource ="Wedstrijd10"
                    GridlineColor =10921638

                    LayoutCachedLeft =7143
                    LayoutCachedTop =5725
                    LayoutCachedWidth =8844
                    LayoutCachedHeight =6040
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5442
                            Top =5725
                            Width =1200
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift34"
                            Caption ="Wedstrijd10"
                            GridlineColor =10921638
                            LayoutCachedLeft =5442
                            LayoutCachedTop =5725
                            LayoutCachedWidth =6642
                            LayoutCachedHeight =6040
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =7029
                    Top =1190
                    Height =315
                    ColumnWidth =2160
                    ColumnOrder =4
                    TabIndex =18
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="TeamID"
                    ControlSource ="TeamID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblTeams.id, tblTeams.TeamNaam FROM tblTeams; "
                    ColumnWidths ="0;2835"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =7029
                    LayoutCachedTop =1190
                    LayoutCachedWidth =8730
                    LayoutCachedHeight =1505
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5442
                            Top =1247
                            Width =780
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift35"
                            Caption ="TeamID"
                            GridlineColor =10921638
                            LayoutCachedLeft =5442
                            LayoutCachedTop =1247
                            LayoutCachedWidth =6222
                            LayoutCachedHeight =1562
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
