Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7371
    DatasheetFontHeight =11
    ItemSuffix =7
    Right =15795
    Bottom =13680
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0x7c0eb60c8192e540
    End
    RecordSource ="SELECT tblToernooi.ID, tblToernooi.ToernooiNaam FROM tblToernooi; "
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
        Begin Line
            BorderLineStyle =0
            Width =1701
            BorderThemeColorIndex =0
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
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            Width =1701
            Height =1701
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
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
            Height =793
            BackColor =15064278
            Name ="Formulierkoptekst"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Width =3360
                    Height =570
                    FontSize =15
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =-2147483617
                    Name ="lblBerekenScorekaarten"
                    Caption ="Berekening Scorestaten"
                    GridlineColor =10921638
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =570
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =5117
            Name ="Details"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =6066
                    Width =336
                    Height =315
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ID"
                    ControlSource ="ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =6066
                    LayoutCachedWidth =6402
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            Left =5612
                            Width =270
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift4"
                            Caption ="ID"
                            GridlineColor =10921638
                            LayoutCachedLeft =5612
                            LayoutCachedWidth =5882
                            LayoutCachedHeight =315
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =56
                    Top =2552
                    Width =2268
                    Height =568
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnOphalenScores"
                    Caption ="--> Scorestaten"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =56
                    LayoutCachedTop =2552
                    LayoutCachedWidth =2324
                    LayoutCachedHeight =3120
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
                    OverlapFlags =93
                    Left =56
                    Top =3799
                    Width =2268
                    Height =568
                    TabIndex =2
                    ForeColor =4210752
                    Name ="btnHTMLUItslagen"
                    Caption ="--> HTML Uitslagen"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =56
                    LayoutCachedTop =3799
                    LayoutCachedWidth =2324
                    LayoutCachedHeight =4367
                    Alignment =1
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
                    OverlapFlags =87
                    Left =56
                    Top =4366
                    Width =2268
                    Height =568
                    TabIndex =3
                    ForeColor =4210752
                    Name ="btnKruisTabel"
                    Caption ="--> HTML Kruistabel"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =56
                    LayoutCachedTop =4366
                    LayoutCachedWidth =2324
                    LayoutCachedHeight =4934
                    Alignment =1
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
                Begin OptionButton
                    OverlapFlags =85
                    Left =3056
                    Top =1870
                    TabIndex =4
                    BorderColor =10921638
                    Name ="optAlle"
                    GridlineColor =10921638

                    LayoutCachedLeft =3056
                    LayoutCachedTop =1870
                    LayoutCachedWidth =3316
                    LayoutCachedHeight =2110
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =56
                            Top =1814
                            Width =2268
                            Height =284
                            FontWeight =700
                            BorderColor =8355711
                            ForeColor =-2147483617
                            Name ="lblAlle"
                            Caption ="Bereken alle staten"
                            GridlineColor =10921638
                            LayoutCachedLeft =56
                            LayoutCachedTop =1814
                            LayoutCachedWidth =2324
                            LayoutCachedHeight =2098
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3055
                    Top =2211
                    Width =3981
                    Height =315
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"3\";\"2\""
                    Name ="cboKiesTeam"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qrOpstellingTeams.Teamnr, qrOpstellingTeams.TeamNaam FROM qrOpstellingTea"
                        "ms WHERE (((qrOpstellingTeams.SessieID)=lngSessieID())) ORDER BY qrOpstellingTea"
                        "ms.Teamnr; "
                    ColumnWidths ="455;2835"
                    GridlineColor =10921638

                    LayoutCachedLeft =3055
                    LayoutCachedTop =2211
                    LayoutCachedWidth =7036
                    LayoutCachedHeight =2526
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =56
                            Top =2211
                            Width =2268
                            Height =284
                            FontWeight =700
                            BorderColor =8355711
                            ForeColor =-2147483617
                            Name ="lblKiesTeam"
                            Caption ="of kies een team"
                            GridlineColor =10921638
                            LayoutCachedLeft =56
                            LayoutCachedTop =2211
                            LayoutCachedWidth =2324
                            LayoutCachedHeight =2495
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =177
                    Left =6236
                    Top =3968
                    Width =576
                    Height =576
                    TabIndex =6
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

                    LayoutCachedLeft =6236
                    LayoutCachedTop =3968
                    LayoutCachedWidth =6812
                    LayoutCachedHeight =4544
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
                End
                Begin OptionButton
                    OverlapFlags =85
                    Left =6803
                    Top =1020
                    TabIndex =7
                    BorderColor =10921638
                    Name ="optHTML"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6803
                    LayoutCachedTop =1020
                    LayoutCachedWidth =7063
                    LayoutCachedHeight =1260
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =3968
                            Top =963
                            Width =2495
                            Height =284
                            FontWeight =700
                            BorderColor =8355711
                            ForeColor =-2147483617
                            Name ="lblHTML"
                            Caption ="Uitvoer naar HTML"
                            GridlineColor =10921638
                            LayoutCachedLeft =3968
                            LayoutCachedTop =963
                            LayoutCachedWidth =6463
                            LayoutCachedHeight =1247
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin OptionButton
                    OverlapFlags =85
                    Left =6803
                    Top =623
                    TabIndex =8
                    BorderColor =10921638
                    Name ="optExcelZichtbaar"
                    GridlineColor =10921638

                    LayoutCachedLeft =6803
                    LayoutCachedTop =623
                    LayoutCachedWidth =7063
                    LayoutCachedHeight =863
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =3968
                            Top =566
                            Width =2495
                            Height =284
                            FontWeight =700
                            BorderColor =8355711
                            ForeColor =-2147483617
                            Name ="lblExcelZichtbaar"
                            Caption ="Excelblad zichtbaar "
                            GridlineColor =10921638
                            LayoutCachedLeft =3968
                            LayoutCachedTop =566
                            LayoutCachedWidth =6463
                            LayoutCachedHeight =850
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =56
                    Top =113
                    Width =3236
                    Height =1273
                    TabIndex =9
                    BorderColor =10921638
                    Name ="grpUitvoernaar"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    GridlineColor =10921638

                    LayoutCachedLeft =56
                    LayoutCachedTop =113
                    LayoutCachedWidth =3292
                    LayoutCachedHeight =1386
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =67
                            Top =113
                            Width =1245
                            Height =315
                            FontWeight =700
                            BorderColor =8355711
                            ForeColor =-2147483617
                            Name ="lblUitvoerNaar"
                            Caption ="Uitvoer Naar"
                            GridlineColor =10921638
                            LayoutCachedLeft =67
                            LayoutCachedTop =113
                            LayoutCachedWidth =1312
                            LayoutCachedHeight =428
                            BackThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =3032
                            Top =351
                            OptionValue =1
                            BorderColor =10921638
                            Name ="optExcel"
                            GridlineColor =10921638

                            LayoutCachedLeft =3032
                            LayoutCachedTop =351
                            LayoutCachedWidth =3292
                            LayoutCachedHeight =591
                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =1417
                                    Top =323
                                    Width =1134
                                    Height =315
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="Bijschrift39"
                                    Caption ="Excel"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =1417
                                    LayoutCachedTop =323
                                    LayoutCachedWidth =2551
                                    LayoutCachedHeight =638
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =3032
                            Top =681
                            TabIndex =1
                            OptionValue =2
                            BorderColor =10921638
                            Name ="optAccess"
                            GridlineColor =10921638

                            LayoutCachedLeft =3032
                            LayoutCachedTop =681
                            LayoutCachedWidth =3292
                            LayoutCachedHeight =921
                            Begin
                                Begin Label
                                    OverlapFlags =95
                                    Left =1417
                                    Top =653
                                    Width =1134
                                    Height =315
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="Bijschrift41"
                                    Caption ="Intern"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =1417
                                    LayoutCachedTop =653
                                    LayoutCachedWidth =2551
                                    LayoutCachedHeight =968
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =3032
                            Top =1011
                            TabIndex =2
                            OptionValue =3
                            BorderColor =10921638
                            Name ="OptBeiden"
                            GridlineColor =10921638

                            LayoutCachedLeft =3032
                            LayoutCachedTop =1011
                            LayoutCachedWidth =3292
                            LayoutCachedHeight =1251
                            Begin
                                Begin Label
                                    OverlapFlags =215
                                    Left =1417
                                    Top =964
                                    Width =1134
                                    Height =315
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="Bijschrift43"
                                    Caption ="Beiden"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =1417
                                    LayoutCachedTop =964
                                    LayoutCachedWidth =2551
                                    LayoutCachedHeight =1279
                                End
                            End
                        End
                    End
                End
                Begin Label
                    OverlapFlags =87
                    Left =56
                    Top =3459
                    Width =2268
                    Height =340
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =-2147483617
                    Name ="lblExtraHTML"
                    Caption ="Extra html pagina's"
                    GridlineColor =10921638
                    LayoutCachedLeft =56
                    LayoutCachedTop =3459
                    LayoutCachedWidth =2324
                    LayoutCachedHeight =3799
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =2993
                    Top =2664
                    Width =4320
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lbGekozenTeam"
                    GridlineColor =10921638
                    LayoutCachedLeft =2993
                    LayoutCachedTop =2664
                    LayoutCachedWidth =7313
                    LayoutCachedHeight =2979
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =93
                    Top =1530
                    Width =7371
                    BorderColor =-2147483617
                    Name ="Lijn0"
                    GridlineColor =10921638
                    LayoutCachedTop =1530
                    LayoutCachedWidth =7371
                    LayoutCachedHeight =1530
                    BorderThemeColorIndex =-1
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =95
                    Width =7371
                    BorderColor =-2147483617
                    Name ="Lijn1"
                    GridlineColor =10921638
                    LayoutCachedWidth =7371
                    BorderThemeColorIndex =-1
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =93
                    Top =3175
                    Width =7371
                    BorderColor =-2147483617
                    Name ="Lijn2"
                    GridlineColor =10921638
                    LayoutCachedTop =3175
                    LayoutCachedWidth =7371
                    LayoutCachedHeight =3175
                    BorderThemeColorIndex =-1
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =93
                    Top =5102
                    Width =7371
                    BorderColor =-2147483617
                    Name ="Lijn3"
                    GridlineColor =10921638
                    LayoutCachedTop =5102
                    LayoutCachedWidth =7371
                    LayoutCachedHeight =5102
                    BorderThemeColorIndex =-1
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =87
                    Left =7370
                    Width =0
                    Height =5102
                    BorderColor =-2147483617
                    Name ="Lijn4"
                    GridlineColor =10921638
                    LayoutCachedLeft =7370
                    LayoutCachedWidth =7370
                    LayoutCachedHeight =5102
                    BorderThemeColorIndex =-1
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =87
                    Width =0
                    Height =5102
                    BorderColor =-2147483617
                    Name ="Lijn5"
                    GridlineColor =10921638
                    LayoutCachedHeight =5102
                    BorderThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3969
                    Top =1814
                    Width =2495
                    Height =315
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =-2147483617
                    Name ="lblTeamNr"
                    Caption ="Team"
                    GridlineColor =10921638
                    LayoutCachedLeft =3969
                    LayoutCachedTop =1814
                    LayoutCachedWidth =6464
                    LayoutCachedHeight =2129
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="Formuliervoettekst"
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

Dim GekozenTeam  As String
Private Sub btnHTMLUItslagen_Click()
Call HTMLViertalUitslagenIn(Sessienr, lngToernooi, lngSessie)
End Sub



Private Sub btnKruisTabel_Click()
Call HTMLViertalKruistabel(lngToernooi)
End Sub


Private Sub btnOphalenScores_Click()
Dim x
Dim WijID As Long

    If (Me.optAlle = False Or IsNull(Me.optAlle)) And (Not IsNull(Me.cboKiesTeam)) Then
        WijID = DLookup("id", "tblTeams", "[Teamnr] = " & Me.cboKiesTeam & " and [ToernooiID] =" & lngToernooi)
        BerekenAlleStaten = False
        Forms("Start_VT").[subProcess].Form.lblTeamNr.Visible = True
        Forms("Start_VT").[subProcess].Form.lblTeamNr.Caption = "--- " & Me.cboKiesTeam & " ---"
        x = VulScoreKaartInSheet(CInt(Me.cboKiesTeam), Sessienr, 2, lngToernooi, ScorestaatIntern, ScorestaatExcel)
        
        If ScorestaatIntern Then
                DoCmd.OpenForm "frmScorestaat", acNormal, , "[ToernooiID] = " & lngToernooi & " and [id] = " & WijID
        Else
            Set xlApp = CreateObject("Excel.Application")
            xlApp.Application.Visible = True
            xlApp.Application.DisplayAlerts = False

            Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)
            Set MySheet = StartBook.Worksheets(Trim(strSheetName))
            'test workfile
             'Me.cboKiesTabblad.Clear
            MySheet.Activate
        End If
        Forms("Start_VT").[subProcess].Form.lblTeamNr.Visible = False
    Else
        If Me.optAlle = True Then
            BerekenAlleStaten = True
            Call AlleScoreStaten_RESULTS(Sessienr, lngToernooi, lngSessie)
        End If
    End If
    BerekenAlleStaten = False
End Sub



Private Sub btnSluiten_Click()
   If CurrentProject.AllForms("Start_VT").IsLoaded = False Then
        DoCmd.Close
    Else
        DoCmd.BrowseTo acBrowseToForm, "frmBegin"
   End If
End Sub


Private Sub cboKiesTeam_AfterUpdate()
  GekozenTeam = Me.cboKiesTeam.Column(1)
  Me.lbGekozenTeam.Caption = GekozenTeam
End Sub

Private Sub Form_Open(Cancel As Integer)
   Dim intSessienr As Integer
         Dim db As Database
         Dim rs As Recordset
        If lngToernooi = 0 Then
             lngToernooi = 1
             lngSessie = 1
             intSessienr = 1
        End If

         Set db = CurrentDb
         Set rs = db.OpenRecordset("select * from tblSessie where [ToernooiD] = " & lngToernooi & " and  [id]= " & lngSessie & " Order by Sessienr")
         lngSessie = rs!Id
         intSessienr = rs!Sessienr
         rs.Close
         db.Close
    
   
   Call InitAll(lngToernooi, lngSessie)
   Me.cboKiesTeam.Requery
   



     Me.optAlle = False
     GekozenTeam = ""
     Me.lblKiesTeam.Visible = True
     Me.cboKiesTeam.Visible = True
     Me.cboKiesTeam.Enabled = True

     Me.optHTML = True
     Me.btnHTMLUItslagen.Visible = True
     Me.btnKruisTabel.Visible = True
     Me.btnHTMLUItslagen.Enabled = True
     Me.btnKruisTabel.Enabled = True
     intUitvoerNaarHTML = True
     ScorestaatIntern = False
     ScorestaatExcel = True
     Forms("Start_VT").[subProcess].Form.lblTeamNr.Visible = False
     
End Sub



Private Sub grpUitvoernaar_AfterUpdate()
    Select Case Me.grpUitvoernaar.Value
    Case 1
    ScorestaatIntern = False
    ScorestaatExcel = True
    Case 2
    ScorestaatIntern = True
    ScorestaatExcel = False
    Case 3
    ScorestaatIntern = True
    ScorestaatExcel = True
    Case Else
    ScorestaatIntern = False
    ScorestaatExcel = False
    End Select
End Sub

Private Sub optAlle_AfterUpdate()
If Me.optAlle = True Then
     Me.lblKiesTeam.Visible = False
     Me.cboKiesTeam.Visible = False
     Me.cboKiesTeam.Enabled = False
     GekozenTeam = ""
  Else
     Me.lblKiesTeam.Visible = True
     Me.cboKiesTeam.Visible = True
     Me.cboKiesTeam.Enabled = True
 End If
End Sub

Private Sub optExcelZichtbaar_Click()
    If optExcelZichtbaar Then
         intExcelZichtbaar = True
    Else
        intExcelZichtbaar = False
    End If
    
End Sub

Private Sub optHTML_AfterUpdate()
If Me.optHTML = True Then
     Me.btnHTMLUItslagen.Visible = True
     Me.btnKruisTabel.Visible = True
     Me.btnHTMLUItslagen.Enabled = True
     Me.btnKruisTabel.Enabled = True
     Me.lblExtraHTML.Visible = True
     intUitvoerNaarHTML = True
   Else
     Me.btnHTMLUItslagen.Visible = False
     Me.btnKruisTabel.Visible = False
     Me.btnHTMLUItslagen.Enabled = False
     Me.btnKruisTabel.Enabled = False
     Me.lblExtraHTML.Visible = False
     intUitvoerNaarHTML = False
End If

End Sub

Private Sub optHTML_Click()
' indien niet, naar uitslag en kruistabel niet zichtbaar
If Me.optHTML = True Then
     Me.btnHTMLUItslagen.Visible = True
     Me.btnKruisTabel.Visible = True
     Me.btnHTMLUItslagen.Enabled = True
     Me.btnKruisTabel.Enabled = True
     Me.lblExtraHTML.Visible = True
     intUitvoerNaarHTML = True
   Else
     Me.btnHTMLUItslagen.Visible = False
     Me.btnKruisTabel.Visible = False
     Me.btnHTMLUItslagen.Enabled = False
     Me.btnKruisTabel.Enabled = False
     Me.lblExtraHTML.Visible = False
     intUitvoerNaarHTML = False
End If


End Sub

Public Function TestOpstelling()
Dim x, i, j As Integer
Dim db As Database
Dim rs, ts As Recordset
Dim intVorigeTeamnr As Integer
Dim speler As String
Dim Spelergevonden As Integer



Set db = CurrentDb
Set rs = db.OpenRecordset("Select * from tblOpstelling where [ToernooiID] = " & lngToernooi & " and [Sessie] = " & Sessienr)
Set ts = db.OpenRecordset("Select * from tblTeam where [ToernooiID] = " & lngToernooi & " Order By Teamnr")

'rijteller = eerste team dat gevonden is

If rs.BOF And rs.EOF Then
     MsgBox ("Er is nog geen opstelling geimporteerd en/of gemaakt")
     TestOpstelling = False
     Exit Function
End If

rs.MoveFirst
intVorigeTeamnr = 0
Do While Not rs.EOF
        If rs!Teamnr - intVorigeTeamnr > 1 Then
           x = MsgBox("Er ontbreekt een team waarschijnlijk team " & rs!Teamnr - intVorigeTeamnr & " Verder testen (J/N) ", vbYesNo)
           If x <> vbYes Then
             TestOpstelling = False
             Exit Function
           End If
        End If
        intVorigeTeamnr = rs!Teamnr
        For i = 1 To 4
            speler = rs.Fields("speler" & i)
            Set ts = db.OpenRecordset("Select * from tblTeam where [ToernooiID] = " & lngToernooi & " and [Teamnr] = " & rs!Teamnr)
                If ts.BOF And ts.EOF Then
                  MsgBox ("Of de teams zijn nog niet geimporteerd of het teamnr " & rs!Teamnr & " ontbreekt ")
                End If
                Spelergevonden = False
                For j = 1 To 8
                If speler = ts.Fields("speler" & j) Then
                    Spelergevonden = True
                    Exit For
                End If
                Next
            ts.Close
            If Spelergevonden = False Then
                MsgBox ("Speler " & speler & " is waarschijnlijk een invaller ")
            End If
            
        Next
        
    
        
        rs.MoveNext
Loop




 'cel 1 avond
 'cel 2 teamnr
 'cel 3 speler1
 'cel 4 speler2
 'cel 5 speler3
 'cel 6 speler4
 'cel 7 tegenstander 1
 'cel 8 tegenstander 2
 'etc tot tegenstander = 0 of leeg
 
 'test op teamnr of dit al niet is geweest
 'test op speler
 'tegenstander   in de kolom mad niet twee keer het zelfde team voorkomen
 

 
 
 

'tel het aantal teams


'indien niet gevonden melding geen opstelling gemaakt
' daarna test of aantal teams en de tegenstanders per wedstrijd of het ok is

'test schema
'tst uitslagen



End Function
