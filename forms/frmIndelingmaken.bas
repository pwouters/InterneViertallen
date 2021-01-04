Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =13049
    DatasheetFontHeight =11
    ItemSuffix =39
    Right =15870
    Bottom =12240
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0x5639299bb194e540
    End
    RecordSource ="qryToernooiSessie"
    Caption ="frmIndelingmaken"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
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
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
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
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin FormHeader
            Height =1303
            BackColor =15064278
            Name ="Formulierkoptekst"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =57
                    Top =57
                    Width =3690
                    Height =969
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Bijschrift10"
                    Caption ="Indeling voorbereiden"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =3747
                    LayoutCachedHeight =1026
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =3968
                    Width =7260
                    Height =570
                    ColumnWidth =3000
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ToernooiNaam"
                    ControlSource ="ToernooiNaam"
                    GridlineColor =10921638

                    LayoutCachedLeft =3968
                    LayoutCachedWidth =11228
                    LayoutCachedHeight =570
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =3585
                    Left =3971
                    Top =566
                    Width =7251
                    Height =315
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="btnKiesSessie"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryToernooiSessie.ID_Sessie, qryToernooiSessie.ToernooiNaam, qryToernooiS"
                        "essie.Sessienr FROM qryToernooiSessie; "
                    ColumnWidths ="0;3402;1134;1134"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3971
                    LayoutCachedTop =566
                    LayoutCachedWidth =11222
                    LayoutCachedHeight =881
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =7031
            Name ="Details"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2552
                    Top =566
                    Width =2268
                    Height =340
                    ColumnWidth =1050
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Sessienr"
                    ControlSource ="Sessienr"
                    GridlineColor =10921638

                    LayoutCachedLeft =2552
                    LayoutCachedTop =566
                    LayoutCachedWidth =4820
                    LayoutCachedHeight =906
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =566
                            Width =2268
                            Height =340
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Sessienr_Bijschrift"
                            Caption ="Sessienr"
                            GridlineColor =10921638
                            LayoutCachedTop =566
                            LayoutCachedWidth =2268
                            LayoutCachedHeight =906
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2552
                    Top =906
                    Width =2268
                    Height =340
                    ColumnWidth =1050
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="AantalTeams"
                    ControlSource ="AantalTeams"
                    GridlineColor =10921638

                    LayoutCachedLeft =2552
                    LayoutCachedTop =906
                    LayoutCachedWidth =4820
                    LayoutCachedHeight =1246
                    Begin
                        Begin Label
                            OverlapFlags =95
                            Top =906
                            Width =2268
                            Height =340
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="AantalTeams_Bijschrift"
                            Caption ="AantalTeams"
                            GridlineColor =10921638
                            LayoutCachedTop =906
                            LayoutCachedWidth =2268
                            LayoutCachedHeight =1246
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Top =2891
                    Width =11160
                    Height =4140
                    TabIndex =2
                    BorderColor =10921638
                    Name ="sbfrmIndeling"
                    SourceObject ="Form.frmIndeling"
                    LinkChildFields ="SessieID"
                    LinkMasterFields ="ID_Sessie"
                    GridlineColor =10921638

                    LayoutCachedTop =2891
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =7031
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2552
                    Top =1246
                    Width =2268
                    Height =340
                    ColumnWidth =2490
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =3484194
                    Name ="wedstrijdvormID"
                    ControlSource ="wedstrijdvormID"
                    RowSourceType ="Value List"
                    RowSource ="0;\"Robin\";1;\"Zwitsers\";2;\"Deens\";3;\"1-2 3-4\";4;\"1-3 4-2\";5;\"1-4 2-3\""
                        ";6;\"Random\""
                    ColumnWidths ="0;1701"
                    StatusBarText ="robin, zwitsers, deens, eerste ronde 1-2, 3-4, tweede ronde 1-4, 3-2"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =2552
                    LayoutCachedTop =1246
                    LayoutCachedWidth =4820
                    LayoutCachedHeight =1586
                    Begin
                        Begin Label
                            OverlapFlags =87
                            Top =1246
                            Width =2268
                            Height =340
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift13"
                            Caption ="Wedstrijdvorm basis"
                            GridlineColor =10921638
                            LayoutCachedTop =1246
                            LayoutCachedWidth =2268
                            LayoutCachedHeight =1586
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =8785
                    Top =1248
                    Width =2271
                    Height =340
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =3484194
                    Name ="cboWedstrijdVormNu"
                    RowSourceType ="Value List"
                    RowSource ="0;\"Robin\";1;\"Zwitsers\";2;\"Deens\";3;\"1-2 3-4\";4;\"1-3 4-2\";5;\"1-4 2-3\""
                        ";6;\"Random\""
                    ColumnWidths ="0;1441"
                    GridlineColor =10921638

                    LayoutCachedLeft =8785
                    LayoutCachedTop =1248
                    LayoutCachedWidth =11056
                    LayoutCachedHeight =1588
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6236
                            Top =1248
                            Width =2258
                            Height =340
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Wedstrijdvorm Nu_Etiket"
                            Caption ="Wedstrijdvorm"
                            EventProcPrefix ="Wedstrijdvorm_Nu_Etiket"
                            GridlineColor =10921638
                            LayoutCachedLeft =6236
                            LayoutCachedTop =1248
                            LayoutCachedWidth =8494
                            LayoutCachedHeight =1588
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =8783
                    Top =907
                    Width =2271
                    Height =340
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtStandNa"
                    ValidationRule ="<=fn_Gespeeld()"
                    ValidationText ="Moet kleiner of gelijk zijn laatst gespeelde ronde"
                    GridlineColor =10921638

                    LayoutCachedLeft =8783
                    LayoutCachedTop =907
                    LayoutCachedWidth =11054
                    LayoutCachedHeight =1247
                    Begin
                        Begin Label
                            OverlapFlags =87
                            Left =6236
                            Top =907
                            Width =2268
                            Height =340
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblStandNa"
                            Caption ="Stand van Ronde"
                            GridlineColor =10921638
                            LayoutCachedLeft =6236
                            LayoutCachedTop =907
                            LayoutCachedWidth =8504
                            LayoutCachedHeight =1247
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =6236
                    Top =170
                    Width =2010
                    Height =340
                    BorderColor =8355711
                    ForeColor =-2147483617
                    Name ="Bijschrift24"
                    Caption ="Indelen nieuw"
                    GridlineColor =10921638
                    LayoutCachedLeft =6236
                    LayoutCachedTop =170
                    LayoutCachedWidth =8246
                    LayoutCachedHeight =510
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Top =170
                    Width =2715
                    Height =340
                    BorderColor =8355711
                    ForeColor =-2147483617
                    Name ="Bijschrift25"
                    Caption ="Sessie/Avond/Dag gegevens"
                    GridlineColor =10921638
                    LayoutCachedTop =170
                    LayoutCachedWidth =2715
                    LayoutCachedHeight =510
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =11338
                    Top =1076
                    Height =567
                    TabIndex =7
                    ForeColor =4210752
                    Name ="btnIndelingMaken"
                    Caption ="Indeling maken"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =11338
                    LayoutCachedTop =1076
                    LayoutCachedWidth =13039
                    LayoutCachedHeight =1643
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2551
                    Top =1870
                    Width =2271
                    Height =340
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtLaatstgespeeldeWedstrijd"
                    ControlSource ="=fn_Gespeeld()"
                    GridlineColor =10921638

                    LayoutCachedLeft =2551
                    LayoutCachedTop =1870
                    LayoutCachedWidth =4822
                    LayoutCachedHeight =2210
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =1870
                            Width =2260
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift30"
                            Caption ="Laatst Gespeelde ronde"
                            GridlineColor =10921638
                            LayoutCachedTop =1870
                            LayoutCachedWidth =2260
                            LayoutCachedHeight =2185
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =11333
                    Top =510
                    Width =1716
                    Height =340
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblVanTot"
                    Caption ="Van 13 tot 14"
                    GridlineColor =10921638
                    LayoutCachedLeft =11333
                    LayoutCachedTop =510
                    LayoutCachedWidth =13049
                    LayoutCachedHeight =850
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =11338
                    Top =1643
                    Height =567
                    TabIndex =9
                    ForeColor =4210752
                    Name ="btnPubliceren"
                    Caption ="Definitief"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =11338
                    LayoutCachedTop =1643
                    LayoutCachedWidth =13039
                    LayoutCachedHeight =2210
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =177
                    Left =12472
                    Top =3401
                    Width =576
                    Height =576
                    TabIndex =10
                    ForeColor =-2147483630
                    Name ="btnSluiten"
                    Caption ="Knop84"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Formulier sluiten"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000030b0200030a080 ,
                        0x1030a0d01030a0f01028a0f01028a0d01028a0801028a0200000000000000000 ,
                        0x00000000000000000000000000000000000000000030b0600030b0f00038c0ff ,
                        0x0038c0ff0040d0ff0040c0ff0038c0ff0038b0ff1028a0f01028a06000000000 ,
                        0x0000000000000000000000000038b0000030b0b00038c0ff0040d0ff0040d0ff ,
                        0x0040d0ff0040d0ff0040d0ff0040d0ff0040c0ff0038c0ff0030b0ff1028a0b0 ,
                        0x1028a00000000000000000000038b0900038c0ff0040e0ff0040e0ff0040d0ff ,
                        0x0040d0ff0040d0ff0040d0ff0040d0ff0040d0ff0040c0ff0040c0ff0030b0ff ,
                        0x1028a090000000000038c0200038c0ff0040e0ff0040e0ff3068e0ffc0d0f0ff ,
                        0x2050e0ff0040d0ff0040d0ff2050d0ffd0d8f0ff3060d0ff0040c0ff0040c0ff ,
                        0x1028a0ff1028a0200038c0800040d0ff0040e0ff0040e0ffc0d0f0fff0f8f0ff ,
                        0xc0d0f0ff2050e0ff2050d0ffc0d0f0fff0f8f0ffc0d0f0ff0040d0ff0040c0ff ,
                        0x0038b0ff1028a0800038c0d00040e0ff0040e0ff0040e0ff2050e0ffc0d0f0ff ,
                        0xf0f8f0ffc0d0f0ffc0d0f0fff0f8f0ffc0d0f0ff2050d0ff0040d0ff0040d0ff ,
                        0x0038c0ff1028a0e00038c0ff0048f0ff0048e0ff0040e0ff0040e0ff2050e0ff ,
                        0xc0d0f0fff0f8f0fff0f8f0ffc0d0f0ff2050d0ff0040d0ff0040d0ff0040d0ff ,
                        0x0040d0ff1028a0ff0038c0ff0048f0ff0048f0ff0048e0ff0040e0ff2050e0ff ,
                        0xc0d0f0fff0f8f0fff0f8f0ffc0d0f0ff2050e0ff0040d0ff0040d0ff0040d0ff ,
                        0x0040d0ff1030a0f00038c0e00048e0ff0048f0ff0048f0ff2058f0ffc0d0f0ff ,
                        0xf0f8f0ffc0d0f0ffc0d0f0fff0f8f0ffc0d0f0ff2050e0ff0040d0ff0040d0ff ,
                        0x0040c0ff1030a0d00040c0900040e0ff0048f0ff0048f0ffc0d0f0fff0f8f0ff ,
                        0xc0d0f0ff2050e0ff2050e0ffc0d0f0fff0f8f0ffc0d0f0ff0040d0ff0040d0ff ,
                        0x0038c0ff0030a0900040c0300040c0ff0048f0ff0048f0ff3068f0ffc0d0f0ff ,
                        0x2058f0ff0040e0ff0040e0ff2050e0ffc0d0f0ff3068e0ff0040e0ff0040d0ff ,
                        0x0030b0ff0030b030000000000040c0a00040d0ff0048f0ff0048f0ff0048f0ff ,
                        0x0048f0ff0048e0ff0040e0ff0040e0ff0040e0ff0040e0ff0040e0ff0038c0ff ,
                        0x0030b09000000000000000000040d0000040c0c00040d0ff0048f0ff0048f0ff ,
                        0x0048f0ff0048f0ff0048e0ff0040e0ff0040e0ff0040e0ff0038c0ff0030b0c0 ,
                        0x0030b0000000000000000000000000000040d0000040c0900040d0ff0040e0ff ,
                        0x0048e0ff0048f0ff0048f0ff0040e0ff0040d0ff0038c0ff0038b0900030b000 ,
                        0x0000000000000000000000000000000000000000000000000040c0300040c090 ,
                        0x0038c0e00038c0ff0038c0ff0038c0e00038c0900038c0300000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =12472
                    LayoutCachedTop =3401
                    LayoutCachedWidth =13048
                    LayoutCachedHeight =3977
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Gradient =0
                    BackColor =-2147483612
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =11338
                    Top =2210
                    Height =567
                    TabIndex =11
                    ForeColor =4210752
                    Name ="btnVerwijder"
                    Caption ="Verwijder indeling"
                    GridlineColor =10921638

                    LayoutCachedLeft =11338
                    LayoutCachedTop =2210
                    LayoutCachedWidth =13039
                    LayoutCachedHeight =2777
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8787
                    Top =567
                    Width =2268
                    Height =340
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    ColumnInfo ="\"\";\"\";\"4\";\"4\""
                    Name ="txtWedstrijdNr"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tblUitslagen.Wedstrijdnr, tblSessie.id FROM tblSessie INNER JOIN"
                        " tblUitslagen ON tblSessie.id = tblUitslagen.SessieID WHERE (((tblSessie.id)=[Fo"
                        "rmulieren]![frmIndelingmaken]![sbfrmIndeling].[Form]![SessieID])) ORDER BY tblUi"
                        "tslagen.Wedstrijdnr; "
                    ValidationRule =">=fnVan() And <=fnTot()"
                    ValidationText ="Ronde moet in de reeds van de sessie zitten"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="=fnVan()"
                    GridlineColor =10921638

                    LayoutCachedLeft =8787
                    LayoutCachedTop =567
                    LayoutCachedWidth =11055
                    LayoutCachedHeight =907
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6236
                            Top =567
                            Width =2265
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblWedstrijdnr"
                            Caption ="Voor Ronde "
                            GridlineColor =10921638
                            LayoutCachedLeft =6236
                            LayoutCachedTop =567
                            LayoutCachedWidth =8501
                            LayoutCachedHeight =882
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3458
                    Top =113
                    Width =336
                    Height =330
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ID_Toernooi"
                    ControlSource ="ID_Toernooi"
                    GridlineColor =10921638

                    LayoutCachedLeft =3458
                    LayoutCachedTop =113
                    LayoutCachedWidth =3794
                    LayoutCachedHeight =443
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2891
                    Top =56
                    Width =411
                    Height =330
                    TabIndex =13
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ID_Sessie"
                    ControlSource ="ID_Sessie"
                    GridlineColor =10921638

                    LayoutCachedLeft =2891
                    LayoutCachedTop =56
                    LayoutCachedWidth =3302
                    LayoutCachedHeight =386
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
Dim IndelingAlGemaakt As Integer

Private Sub btnIndelingMaken_Click()
Dim TeamID() As Integer
Dim Indeling As Variant
'test of er al een indeling is gemaakt
'

If IsNull(Me.cboWedstrijdVormNu) Or Me.cboWedstrijdVormNu = "" Then
  MsgBox "Geen rekenvorm ingevuld"
  Exit Sub
End If


ReDim TeamID(Me.AANTALTEAMS)

Dim db As Database
Dim rs As Recordset
Dim ts As Recordset

Set db = CurrentDb
Set ts = db.OpenRecordset("Select * from tblTeams where ToernooiID = " & Me.ID_Toernooi & " Order by Teamnr;")
ts.MoveFirst
Do While Not ts.EOF
    TeamID(ts!Teamnr) = ts!id
    ts.MoveNext
Loop
ts.Close


Set rs = db.OpenRecordset("select * from tblIndeling where Wedstrijdnr = " & txtWedstrijdNr & " and ToernooiID = " & Me.ID_Toernooi)
If Not (rs.BOF And rs.EOF) Then
    question = MsgBox("Reeds een indeling gemaakt in tblIndeling Overschrijven (J/N) ", vbYesNo)
    If question = vbNo Then
        rs.Close
        db.Close
        Exit Sub
    End If
    'verwijder indeling
    Dim qd As QueryDef
    sql = "Delete from tblIndeling where Wedstrijdnr = " & txtWedstrijdNr & " and ToernooiID = " & Me.ID_Toernooi
    Set qd = db.CreateQueryDef("", sql)
    qd.Execute
    Set qd = Nothing
End If

'0;"Robin";1;"Zwitsers";2;"Deens";3;"1-2 3-4";4;"1-3 4-2";5;"1-4 2-3";6;"Random"
Select Case Me.cboWedstrijdVormNu
Case 0
    'laden vanuit intern bestand
    Set rs = db.OpenRecordset("select * from tblSchema where Wedstrijdronde = " & txtWedstrijdNr & " and ToernooiID = " & Me.ID_Toernooi)
    If rs.BOF And rs.EOF Then
        rs.Close
        db.Close
        MsgBox ("Schema is niet opgeslagen ")
        Exit Sub
    End If
    ReDim Indeling(Me.AANTALTEAMS \ 2, 1)
    i = 0
    Do While Not rs.EOF
        i = i + 1
        Indeling(i, 0) = TeamID(rs!TeamThuis)
        Indeling(i, 1) = TeamID(rs!TeamUit)
        rs.MoveNext
    Loop
    rs.Close
Case 1
'Zwitsers
Indeling = BepaalZwitsersIntern(Me.AANTALTEAMS, Me.txtStandNa, Me.txtLaatstgespeeldeWedstrijd, Me.ID_Toernooi, 16)
'parameters
Case 2
Indeling = BepaalDeensIntern(Me.AANTALTEAMS, Me.txtLaatstgespeeldeWedstrijd, Me.ID_Toernooi, 16)
Case 3
Indeling = BepaalIndeling1234(Me.AANTALTEAMS, Me.ID_Toernooi, 16)
Case 4
Indeling = BepaalIndeling1342(Me.AANTALTEAMS, Me.ID_Toernooi, 16)
Case 5
Indeling = BepaalIndeling1423(Me.AANTALTEAMS, Me.ID_Toernooi, 16)
Case 6
Indeling = RekenRandomIntern(Me.AANTALTEAMS, Me.txtLaatstgespeeldeWedstrijd, Me.ID_Toernooi, 16)
Case Else
  MsgBox "Geen rekenvorm ingevuld"
  Exit Sub
End Select


Set ts = db.OpenRecordset("tblIndeling")
For i = 1 To Me.AANTALTEAMS \ 2
    ts.AddNew
    ts!TeamThuisID = TeamID(Indeling(i, 0))
    ts!TeamUitID = TeamID(Indeling(i, 1))
    ts!Wedstrijdnr = Me.txtWedstrijdNr
    ts!SessieID = Me.ID_Sessie
    ts!ToernooiID = Me.ID_Toernooi
    ts!WedstrijdVorm = Me.cboWedstrijdVormNu
    ts!StandNaWedstrijdNr = Me.txtStandNa
    ts.Update
Next
ts.Close
db.Close

Me.sbfrmIndeling.Requery
IndelingAlGemaakt = True
Me.btnVerwijder.Visible = IndelingAlGemaakt
Me.btnVerwijder.Enabled = IndelingAlGemaakt
Me.btnPubliceren.Visible = IndelingAlGemaakt
Me.btnPubliceren.Enabled = IndelingAlGemaakt
'genereer indeling
'voeg toe aan tabel
'fris subform op

End Sub

Private Sub btnKiesSessie_AfterUpdate()
    Dim rs As Recordset
    Dim x
    Dim Criterium As String
    Set rs = Recordset
   
    Criterium = "[ID_Sessie]= " & btnKiesSessie
    rs.FindFirst Criterium
    
    If rs.NoMatch Then
        MsgBox "Geen Sessie bekend in de database"
        btnKiesSessie = ""
        Exit Sub
    Else
        Me.Bookmark = rs.Bookmark
        btnKiesSessie = ""
    End If
    
    'wedstrijd in sessie van tot
    WedstrijdenVanaf = GespeeldeWedstrijden_vanaf(Me.ID_Sessie, Me.ID_Toernooi)
    WedstrijdenTot = GespeeldeWedstrijden_tot(Me.ID_Sessie, Me.ID_Toernooi)
    GespeeldeRondes = DMax("WedstrijdNr", "qryGespeeldeWedstrijden", "ToernooiID = " & Me.ID_Toernooi)
    Me.lblVanTot.Caption = "Van " & WedstrijdenVanaf & " tot " & WedstrijdenTot
     
    Me.txtWedstrijdNr = fnVan()
     'Me.txtWedstrijdNr = Me.Wedstrijdnr.Value
    
    Me.txtStandNa = fn_Gespeeld()
    If Not IsNull(DLookup("Wedstrijdnr", "tblIndeling", "[sessieiD] = " & Me.ID_Sessie & " and [WedstrijdNr] = " & Me.txtWedstrijdNr)) Then
        IndelingAlGemaakt = True
        Else
        IndelingAlGemaakt = False
    End If
     Me.txtWedstrijdNr.Requery
    [sbfrmIndeling].Form.Filter = "[WedstrijdNr] = " & Me.txtWedstrijdNr
    [sbfrmIndeling].Form.FilterOn = True
    [sbfrmIndeling].Requery
    Me.btnVerwijder.Visible = IndelingAlGemaakt
    Me.btnVerwijder.Enabled = IndelingAlGemaakt
    Me.btnPubliceren.Visible = IndelingAlGemaakt
    Me.btnPubliceren.Enabled = IndelingAlGemaakt
   

End Sub

Private Sub btnPubliceren_Click()
Dim question As Integer
Dim i, j As Integer
Dim db As Database

Dim rs_Indeling As Recordset
Dim rs_Schema As Recordset
Dim rs_Uitslagen As Recordset

Dim qd As QueryDef

'intern

Dim sql_Indeling     As String
Dim sql_Schema       As String
Dim sql_Uitslagen    As String

Dim tbl_Indeling     As String
Dim tbl_Schema       As String
Dim tbl_Uitslagen    As String


Dim del_Schema       As String
Dim del_Uitslagen    As String


Dim int_Indeling As Integer

Dim int_Schema_append As Integer
Dim int_Schema_mutate As Integer

Dim int_Uitslagen_append As Integer
Dim int_Uitslagen_Mutate As Integer


sql_Indeling = ""
sql_Schema = ""
sql_Uitslagen = ""
Set db = CurrentDb


tbl_Indeling = "select * from tblIndeling where "
tbl_Schema = "select * from tblSchema where "
tbl_Uitslagen = "select * from tblUitslagen where "


del_Schema = "delete from tblSchema where "
del_Uitslagen = "delete * from tblUitslagen where "

'
'muteren/toevoegen schema
sql_Indeling = tbl_Indeling & " Wedstrijdnr = " & txtWedstrijdNr & " and ToernooiID = " & Me.ID_Toernooi
Set rs_Indeling = db.OpenRecordset(sql_Indeling)

If rs_Indeling.BOF And rs_Indeling.EOF Then
    MsgBox "Zijn geen records aangemaakt", , "Opslaan Indeling"
    rs_Indeling.Close
    db.Close
    Exit Sub
End If

int_Indeling = True


sql_Schema = tbl_Schema & " Wedstrijdronde = " & txtWedstrijdNr & " and ToernooiID = " & Me.ID_Toernooi
Set rs_Schema = db.OpenRecordset(sql_Schema)

If Not (rs_Schema.BOF And rs_Schema.EOF) Then
    question = MsgBox("Ronde is reeds in het schema opgenomen, Overschrijden", vbYesNo, "Opslaan Indeling")
    If question = vbNo Then
        rs_Indeling.Close
        rs_Schema.Close
        db.Close
        Exit Sub
    End If
    int_Schema_mutate = True
    int_Schema_append = True
  Else
    question = MsgBox("Ronde is nog niet in het schema opgenomen, Toevoegen", vbYesNo, "Opslaan Indeling")
    If question = vbNo Then
        rs_Indeling.Close
        rs_Schema.Close
        db.Close
        Exit Sub
    End If
    int_Schema_mutate = False
    int_Schema_append = True
End If

sql_Uitslagen = tbl_Uitslagen & " Wedstrijdnr = " & txtWedstrijdNr & " and ToernooiID = " & Me.ID_Toernooi
Set rs_Indeling = db.OpenRecordset(sql_Uitslagen)

If Not (rs_Uitslagen.BOF And rs_Uitslagen.EOF) Then
    question = MsgBox("Rondegegevens zijn reeds in de uitslagen opgenomen, Overschrijden", vbYesNo, "Opslaan Indeling")
    If question = vbNo Then
        rs_Indeling.Close
        rs_Uitslagen.Close
        db.Close
        Exit Sub
    End If
    int_Uitslagen_Mutate = True
    int_Uitslagen_append = True
  Else
    question = MsgBox("Rondegegevens zijn nog niet in de uitslagen opgenomen, Toevoegen", vbYesNo, "Opslaan Indeling")
    If question = vbNo Then
        rs_Indeling.Close
        rs_Uitslagen.Close
        db.Close
        Exit Sub
    End If
    int_Uitslagen_Mutate = False
    int_Uitslagen_append = True
End If

'muteren  (verwijderen + toevoegen/toevoegen teamuitslagen

If int_Schema_mutate Then
    sql_Schema = del_Schema & " Wedstrijdronde = " & txtWedstrijdNr & " and ToernooiID = " & Me.ID_Toernooi
    Set qd = db.CreateQueryDef("", sql_Schema)
    qd.Execute
    Set qd = Nothing
End If

If int_Uitslagen_Mutate Then
    sql_Uitslagen = del_Uitslagen & " Wedstrijdronde = " & txtWedstrijdNr & " and ToernooiID = " & Me.ID_Toernooi
    Set qd = db.CreateQueryDef("", sql_Uitslagen)
    qd.Execute
    Set qd = Nothing
End If

If int_Schema_append Then
    i = 1
    rs_Indeling.MoveFirst
    Do While Not rs_Indeling.EOF
        rs_Schema.AddNew
        rs_Schema!ToernooiID = rs_Indeling!ToernooiID
        rs_Schema!SessieID = rs_Indeling!SessieID
        rs_Schema!Wedstrijdronde = rs_Indeling!Wedstrijdnr
        rs_Schema!Paring = i
        rs_Schema!TeamThuis = rs_Indeling!TeamThuisID
        rs_Schema!TeamUit = rs_Indeling!TeamUitID
        rs_Schema.Update
        i = i + 1
        rs_Indeling.MoveNext
    Loop
End If


If int_Uitslagen_append Then
    i = 1
    rs_Indeling.MoveFirst
    Do While Not rs_Indeling.EOF
        rs_Uitslagen.AddNew
        rs_Uitslagen!ToernooiID = rs_Indeling!ToernooiID
        rs_Uitslagen!SessieID = rs_Indeling!SessieID
        rs_Uitslagen!Wedstrijdnr = rs_Indeling!Wedstrijdnr
        rs_Uitslagen!TeamIDThuis = rs_Indeling!TeamThuisID
        rs_Uitslagen!TeamIDUit = rs_Indeling!TeamUitID
        rs_Uitslagen.Update
        i = i + 1
        rs_Indeling.MoveNext
    Loop
End If
    
rs_Uitslagen.Close
rs_Schema.Close
rs_Indeling.Close

End Sub

Private Sub btnSluiten_Click()
DoCmd.Close
End Sub

Private Sub Form_Open(Cancel As Integer)
    
    If CurrentProject.AllForms("Start_VT").IsLoaded = True Then
         Dim rs As Recordset
         Dim x
         Dim Criterium As String
         Set rs = Recordset
        
         Criterium = "[ID_Sessie]= " & lngSessie
         rs.FindFirst Criterium
         
         If rs.NoMatch Then
             MsgBox "Geen Sessie bekend in de database"
             btnKiesSessie = ""
             Exit Sub
         Else
             Me.Bookmark = rs.Bookmark
             btnKiesSessie = ""
         End If
         Me.txtWedstrijdNr = intWedstrijdronde
    End If

    WedstrijdenVanaf = GespeeldeWedstrijden_vanaf(Me.ID_Sessie, Me.ID_Toernooi)
    WedstrijdenTot = GespeeldeWedstrijden_tot(Me.ID_Sessie, Me.ID_Toernooi)
    GespeeldeRondes = DMax("WedstrijdNr", "qryGespeeldeWedstrijden", "ToernooiID = " & Me.ID_Toernooi)
    Me.lblVanTot.Caption = "Van " & WedstrijdenVanaf & " tot " & WedstrijdenTot
    
    If CurrentProject.AllForms("frmSchemaVerwerking").IsLoaded = False Then
        Me.txtWedstrijdNr = fnVan()
    End If
    'Me.txtWedstrijdNr = Me!Wedstrijdnr.Value
   
    Me.txtStandNa = fn_Gespeeld()
    If Not IsNull(DLookup("Wedstrijdnr", "tblIndeling", "[sessieID] = " & Me.ID_Sessie & " and [WedstrijdNr] = " & Me.txtWedstrijdNr)) Then
        IndelingAlGemaakt = True
    Else
        IndelingAlGemaakt = False
    End If
    [sbfrmIndeling].Form.Filter = "[WedstrijdNr] = " & Me.txtWedstrijdNr
    [sbfrmIndeling].Form.FilterOn = True
    [sbfrmIndeling].Requery
    Me.btnVerwijder.Visible = IndelingAlGemaakt
    Me.btnVerwijder.Enabled = IndelingAlGemaakt
    Me.btnPubliceren.Visible = IndelingAlGemaakt
    Me.btnPubliceren.Enabled = IndelingAlGemaakt
End Sub

Private Sub txtWedstrijdNr_AfterUpdate()
    If Not IsNull(DLookup("Wedstrijdnr", "tblIndeling", "[sessieiD] = " & Me.ID_Sessie & " and [WedstrijdNr] = " & Me.txtWedstrijdNr)) Then
        IndelingAlGemaakt = True
    Else
        IndelingAlGemaakt = False
    End If
    Me.btnVerwijder.Visible = IndelingAlGemaakt
    Me.btnVerwijder.Enabled = IndelingAlGemaakt
    Me.btnPubliceren.Visible = IndelingAlGemaakt
    Me.btnPubliceren.Enabled = IndelingAlGemaakt
    [sbfrmIndeling].Form.Filter = "[WedstrijdNr] = " & Me.txtWedstrijdNr
    [sbfrmIndeling].Form.FilterOn = True
    [sbfrmIndeling].Requery
End Sub
