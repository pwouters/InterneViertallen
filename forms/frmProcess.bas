Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9231
    DatasheetFontHeight =11
    ItemSuffix =32
    Left =2520
    Top =1125
    Right =17445
    Bottom =10995
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0x6f1d1de43490e540
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
        Begin ListBox
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
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
        Begin EmptyCell
            Height =240
            GridlineColor =12632256
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =6804
            Name ="Details"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =570
                    Top =345
                    Width =6120
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Bijschrift0"
                    Caption ="Process  Viertallen"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =570
                    LayoutCachedTop =345
                    LayoutCachedWidth =6690
                    LayoutCachedHeight =660
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =3403
                    Top =964
                    Width =4821
                    Height =315
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="cboKiesToernooi"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [tblToernooi].[ID], [tblToernooi].[ToernooiNaam] FROM tblToernooi ORDER B"
                        "Y [ToernooiNaam]; "
                    ColumnWidths ="0;1440"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3403
                    LayoutCachedTop =964
                    LayoutCachedWidth =8224
                    LayoutCachedHeight =1279
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =577
                            Top =964
                            Width =2640
                            Height =320
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Kies Toernooi_Etiket"
                            Caption ="Kies Toernooi"
                            EventProcPrefix ="Kies_Toernooi_Etiket"
                            GridlineColor =10921638
                            LayoutCachedLeft =577
                            LayoutCachedTop =964
                            LayoutCachedWidth =3217
                            LayoutCachedHeight =1284
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7426
                    Top =113
                    Height =315
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ID"
                    ControlSource ="ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =7426
                    LayoutCachedTop =113
                    LayoutCachedWidth =9127
                    LayoutCachedHeight =428
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =215
                            Left =5725
                            Top =113
                            Width =270
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift4"
                            Caption ="ID"
                            GridlineColor =10921638
                            LayoutCachedLeft =5725
                            LayoutCachedTop =113
                            LayoutCachedWidth =5995
                            LayoutCachedHeight =428
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =566
                    Top =2834
                    Width =2091
                    Height =568
                    TabIndex =2
                    ForeColor =4210752
                    Name ="btnOphalenScores"
                    Caption ="--> Scorestaten"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =2834
                    LayoutCachedWidth =2657
                    LayoutCachedHeight =3402
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
                    OverlapFlags =95
                    Left =566
                    Top =3401
                    Width =2091
                    Height =568
                    TabIndex =3
                    ForeColor =4210752
                    Name ="btnHTMLUItslagen"
                    Caption ="--> HTML Uitslagen"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =3401
                    LayoutCachedWidth =2657
                    LayoutCachedHeight =3969
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
                    Left =566
                    Top =3968
                    Width =2076
                    Height =568
                    TabIndex =4
                    ForeColor =4210752
                    Name ="btnKruisTabel"
                    Caption ="--> HTML Kruistabel"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =3968
                    LayoutCachedWidth =2642
                    LayoutCachedHeight =4536
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
                    Left =3283
                    Top =2978
                    TabIndex =5
                    BorderColor =10921638
                    Name ="optAlle"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3283
                    LayoutCachedTop =2978
                    LayoutCachedWidth =3543
                    LayoutCachedHeight =3218
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =3623
                            Top =2948
                            Width =555
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblAlle"
                            Caption ="alle"
                            GridlineColor =10921638
                            LayoutCachedLeft =3623
                            LayoutCachedTop =2948
                            LayoutCachedWidth =4178
                            LayoutCachedHeight =3263
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6345
                    Top =2948
                    Width =2886
                    Height =315
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"3\";\"2\""
                    Name ="cboKiesTeam"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryTeams_per_Toernooi.Teamnr, qryTeams_per_Toernooi.TeamNaam FROM qryTeam"
                        "s_per_Toernooi WHERE (((qryTeams_per_Toernooi.ID)=lngToernooiID())) ORDER BY qry"
                        "Teams_per_Toernooi.Teamnr; "
                    ColumnWidths ="454;2835"
                    GridlineColor =10921638

                    LayoutCachedLeft =6345
                    LayoutCachedTop =2948
                    LayoutCachedWidth =9231
                    LayoutCachedHeight =3263
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5199
                            Top =2948
                            Width =975
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblKiesTeam"
                            Caption ="Team"
                            GridlineColor =10921638
                            LayoutCachedLeft =5199
                            LayoutCachedTop =2948
                            LayoutCachedWidth =6174
                            LayoutCachedHeight =3263
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2760
                    Left =3403
                    Top =1360
                    Width =4881
                    Height =315
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"3\";\"2\""
                    Name ="cboKiesSessie"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblSessie.Sessienr, tblSessie.Sessienaam FROM tblSessie WHERE (((tblSessi"
                        "e.ToernooID)=lngToernooiID())) ORDER BY tblSessie.Sessienr, tblSessie.Sessienaam"
                        "; "
                    ColumnWidths ="567;2835"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    GridlineColor =10921638

                    LayoutCachedLeft =3403
                    LayoutCachedTop =1360
                    LayoutCachedWidth =8284
                    LayoutCachedHeight =1675
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =577
                            Top =1360
                            Width =2655
                            Height =320
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Sessienr_Etiket"
                            Caption ="Sessienr"
                            GridlineColor =10921638
                            LayoutCachedLeft =577
                            LayoutCachedTop =1360
                            LayoutCachedWidth =3232
                            LayoutCachedHeight =1680
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =566
                    Top =5102
                    Width =2091
                    Height =568
                    TabIndex =8
                    ForeColor =4210752
                    Name ="btnImportTeams"
                    Caption ="Import Teams"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =5102
                    LayoutCachedWidth =2657
                    LayoutCachedHeight =5670
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
                    OverlapFlags =95
                    Left =566
                    Top =5669
                    Width =2091
                    Height =568
                    TabIndex =9
                    ForeColor =4210752
                    Name ="btnImportOpstelling"
                    Caption ="Import Opstelling"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =5669
                    LayoutCachedWidth =2657
                    LayoutCachedHeight =6237
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
                    Visible = NotDefault
                    OverlapFlags =87
                    Left =566
                    Top =6236
                    Width =2091
                    Height =568
                    TabIndex =10
                    ForeColor =4210752
                    Name ="btnImportIndeling"
                    Caption ="Import Indeling"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =6236
                    LayoutCachedWidth =2657
                    LayoutCachedHeight =6804
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
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =761
                    Top =4705
                    Width =1695
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblImport"
                    Caption ="Van werkbestand"
                    GridlineColor =10921638
                    LayoutCachedLeft =761
                    LayoutCachedTop =4705
                    LayoutCachedWidth =2456
                    LayoutCachedHeight =5020
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =2834
                    Top =5102
                    Width =2091
                    Height =568
                    TabIndex =11
                    ForeColor =4210752
                    Name ="btnExportTeams"
                    Caption ="Export Teams"
                    GridlineColor =10921638

                    LayoutCachedLeft =2834
                    LayoutCachedTop =5102
                    LayoutCachedWidth =4925
                    LayoutCachedHeight =5670
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
                    Visible = NotDefault
                    OverlapFlags =95
                    Left =2834
                    Top =5669
                    Width =2091
                    Height =568
                    TabIndex =12
                    ForeColor =4210752
                    Name ="btnExportOpstelling"
                    Caption ="Export Opstelling"
                    GridlineColor =10921638

                    LayoutCachedLeft =2834
                    LayoutCachedTop =5669
                    LayoutCachedWidth =4925
                    LayoutCachedHeight =6237
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
                    Visible = NotDefault
                    OverlapFlags =87
                    Left =2834
                    Top =6236
                    Width =2091
                    Height =568
                    TabIndex =13
                    ForeColor =4210752
                    Name ="btnExportIndeling"
                    Caption ="Export Indeling"
                    GridlineColor =10921638

                    LayoutCachedLeft =2834
                    LayoutCachedTop =6236
                    LayoutCachedWidth =4925
                    LayoutCachedHeight =6804
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
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =2991
                    Top =4705
                    Width =1770
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Bijschrift25"
                    Caption ="Naar werkbestand"
                    GridlineColor =10921638
                    LayoutCachedLeft =2991
                    LayoutCachedTop =4705
                    LayoutCachedWidth =4761
                    LayoutCachedHeight =5020
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =177
                    Left =7937
                    Top =5102
                    Width =576
                    Height =576
                    TabIndex =14
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

                    LayoutCachedLeft =7937
                    LayoutCachedTop =5102
                    LayoutCachedWidth =8513
                    LayoutCachedHeight =5678
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
                Begin CommandButton
                    OverlapFlags =93
                    Left =5102
                    Top =5102
                    Width =2106
                    Height =568
                    TabIndex =15
                    ForeColor =4210752
                    Name ="btnNewToernooi"
                    Caption ="Nieuw Toernooi"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =5102
                    LayoutCachedTop =5102
                    LayoutCachedWidth =7208
                    LayoutCachedHeight =5670
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
                    Left =5102
                    Top =5669
                    Width =2106
                    Height =568
                    TabIndex =16
                    ForeColor =4210752
                    Name ="btnNieuweSessie"
                    Caption ="Nieuwe Sessie"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =5102
                    LayoutCachedTop =5669
                    LayoutCachedWidth =7208
                    LayoutCachedHeight =6237
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
                    Left =3288
                    Top =2324
                    TabIndex =17
                    BorderColor =10921638
                    Name ="optHTML"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3288
                    LayoutCachedTop =2324
                    LayoutCachedWidth =3548
                    LayoutCachedHeight =2564
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =565
                            Top =2267
                            Width =2085
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblHTML"
                            Caption ="Uitvoer naar HTML"
                            GridlineColor =10921638
                            LayoutCachedLeft =565
                            LayoutCachedTop =2267
                            LayoutCachedWidth =2650
                            LayoutCachedHeight =2582
                        End
                    End
                End
                Begin OptionButton
                    OverlapFlags =85
                    Left =6803
                    Top =2267
                    TabIndex =18
                    BorderColor =10921638
                    Name ="optExcelZichtbaar"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6803
                    LayoutCachedTop =2267
                    LayoutCachedWidth =7063
                    LayoutCachedHeight =2507
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4535
                            Top =2267
                            Width =2100
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblExcelZichtbaar"
                            Caption ="Excel Zichtbaar"
                            GridlineColor =10921638
                            LayoutCachedLeft =4535
                            LayoutCachedTop =2267
                            LayoutCachedWidth =6635
                            LayoutCachedHeight =2582
                        End
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

Private Sub btnHTMLUItslagen_Click()
Call HTMLViertalUitslagenIn(CInt(Me.cboKiesSessie), CInt(Me.cboKiesToernooi), lngSessie)
End Sub

Private Sub btnImportOpstelling_Click()
Dim rijteller As Integer
Dim db As Database
Dim rs As Recordset
Dim MySheet As Worksheet
Dim StartBook As Workbook
Dim strWorkFile As String
Dim question As Integer
Dim Sessienr As Integer
Dim sessieaanwezig As Integer

Sessienr = CInt(Me.cboKiesSessie)

Set db = CurrentDb
Set rs = db.OpenRecordset("select * from tblOpstelling where [ToernooiID] = " & lngToernooi & " and [Sessie] = " & CInt(Sessienr))
If Not (rs.BOF And rs.EOF) Then
    MsgBox ("Er is  reeds opstelling geladen of aanwezig")
    rs.Close
    db.Close
    Exit Sub
End If

'test of er een werkbestand is
strWorkFile = WORKFOLDER & WORKFILE

If Not fnExists(strWorkFile) Then
     MsgBox ("Er is nog geen excel bestand aangemaakt")
     Exit Sub
End If

     
Set xlApp = CreateObject("Excel.Application")
xlApp.Application.Visible = True
xlApp.Application.DisplayAlerts = False
Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)
Set MySheet = StartBook.Worksheets("Import_Opstelling")

If MySheet.Cells(2, 1) = "" Then
    question = MsgBox("Er zijn geen opstelling aanwezig")
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    Exit Sub
End If
    sessieaanwezig = False
 rijteller = 2
  Do While MySheet.Cells(rijteller, 1) <> ""
    If MySheet.Cells(rijteller, 1).Value = Sessienr Then
             If Not sessieaanwezig Then sessieaanwezig = True
            rs.AddNew
            rs!ToernooiID = lngToernooi
            rs!Sessie = MySheet.Cells(rijteller, 1).Value
            rs!TeamNr = MySheet.Cells(rijteller, 2).Value
            rs!Speler1 = MySheet.Cells(rijteller, 3).Value
            rs!Speler2 = MySheet.Cells(rijteller, 4).Value
            rs!Speler3 = MySheet.Cells(rijteller, 5).Value
            rs!Speler4 = MySheet.Cells(rijteller, 6).Value
            Kolom = 7
            intWedstrijd = 1
             Do While MySheet.Cells(rijteller, Kolom).Value <> "" And intWedstrijd < 12
                rs.Fields("wedstrijd" & intWedstrijd) = MySheet.Cells(rijteller, Kolom).Value
                Kolom = Kolom + 1
                intWedstrijd = intWedstrijd + 1
             Loop
            rs.Update
     End If
    rijteller = rijteller + 1
   Loop
    
    rs.Close
    db.Close
  
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    
  
   If Not sessieaanwezig Then
    question = MsgBox("Er is geen opstelling aanwezig in de excel file")
  End If
 End Sub

Private Sub btnImportTeams_Click()
  'kijk eerst op er al teams opgenomen zijn
  
Dim db As Database
Dim rs As Recordset
Dim MySheet As Worksheet
Dim StartBook As Workbook
Dim strWorkFile As String
Dim question As Integer
Set db = CurrentDb
Set rs = db.OpenRecordset("select * from tblTeams where [ToernooiID] = " & lngToernooi)
If Not (rs.BOF And rs.EOF) Then
    MsgBox ("Er zijn reeds teams geladen of aanwezig")
    rs.Close
    db.Close
    
    Exit Sub
End If

'test of er een werkbestand is
strWorkFile = WORKFOLDER & WORKFILE

If Not fnExists(strWorkFile) Then
     MsgBox ("Er is nog geen excel bestand aangemaakt")
     Exit Sub
End If

     
Set xlApp = CreateObject("Excel.Application")
xlApp.Application.Visible = True
xlApp.Application.DisplayAlerts = False
Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)
Set MySheet = StartBook.Worksheets("Teams")

If MySheet.Cells(2, 1) = "" Then
    question = MsgBox("Er zijn geen teams aanwezig")
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    Exit Sub
End If
    
 rijteller = 2
 
 Do While MySheet.Cells(rijteller, 1) <> ""
    rs.AddNew
    rs!ToernooiID = lngToernooi
    rs!TeamNr = MySheet.Cells(rijteller, 1).Value
    rs!TeamNaam = MySheet.Cells(rijteller, 2).Value
    'niet meer dan 8 spelers
    Kolom = 3
    speler = 1
    Do While MySheet.Cells(rijteller, Kolom).Value <> "" And speler < 9
       rs.Fields("speler" & speler) = MySheet.Cells(rijteller, Kolom).Value
       Kolom = Kolom + 1
       speler = speler + 1
    Loop
    rs.Update
    rijteller = rijteller + 1
 Loop
 rs.Close
 db.Close
 
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
 
End Sub

Private Sub btnKruisTabel_Click()
Call HTMLViertalKruistabel(CInt(Me.cboKiesToernooi))
End Sub

Private Sub btnNewToernooi_Click()
' Kies worktemplate

'' strTemplate_Folder


' indien ja opslaan als
' add record tblTooernooi
' add record sessie gegevens (sessie no 1)
End Sub

Private Sub btnNieuweSessie_Click()
    'hoog sessie nr op
    'dupliceer gegevens van de vorige sessie
End Sub

Private Sub btnOphalenScores_Click()
Dim X
        
    If (Me.optAlle = False Or IsNull(Me.optAlle)) And (Not IsNull(Me.cboKiesTeam)) Then
        X = VulScoreKaartInSheet(CInt(Me.cboKiesTeam), CInt(Me.cboKiesSessie), 2, CInt(Me.cboKiesToernooi))
    Else
        Call AlleScoreStaten_RESULTS(CInt(Me.cboKiesSessie), CInt(Me.cboKiesToernooi), lngSessie)
    End If


End Sub



Private Sub btnSluiten_Click()
    DoCmd.Close
End Sub

Private Sub cboKiesSessie_AfterUpdate()
     lngSessie = DLookup("id", "tblSessie", "Sessienr=" & Me.cboKiesSessie & " and ToernooiD = " & lngToernooi)
     Call InitAll(lngToernooi, lngSessie)
  
End Sub

Private Sub cboKiesToernooi_AfterUpdate()
    lngToernooi = Me.cboKiesToernooi
    'lngSessie = DLookup("id", "tblSessie", "Sessienr=" & Me.cboKiesSessie & " and ToernooiD = " & lngToernooi)
     lngSessie = DLookup("id", "tblSessie", "Sessienr=" & 1 & " and ToernooiD = " & lngToernooi)
    
     Call InitAll(lngToernooi, lngSessie)
     Me.cboKiesTeam.Requery
     Me.cboKiesSessie.Requery
     Me.cboKiesSessie = Sessienr
   
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
   Me.cboKiesSessie.Requery
   Me.cboKiesSessie.Value = lngSessie
   Me.cboKiesToernooi.Value = lngToernooi



     Me.optAlle = False
     Me.lblKiesTeam.Visible = True
     Me.cboKiesTeam.Visible = True
     Me.cboKiesTeam.Enabled = True

     Me.optHTML = True
     Me.btnHTMLUItslagen.Visible = True
     Me.btnKruisTabel.Visible = True
     Me.btnHTMLUItslagen.Enabled = True
     Me.btnKruisTabel.Enabled = True
     intUitvoerNaarHTML = True


End Sub



Private Sub optAlle_AfterUpdate()
If Me.optAlle = True Then
     Me.lblKiesTeam.Visible = False
     Me.cboKiesTeam.Visible = False
     Me.cboKiesTeam.Enabled = False
  Else
     Me.lblKiesTeam.Visible = True
     Me.cboKiesTeam.Visible = True
     Me.cboKiesTeam.Enabled = True
 End If
End Sub

Private Sub optExcelZichtbaar_Click()
    If optExcelZichtbaar Then
         intUitvoerNaarHTML = True
    Else
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
     intUitvoerNaarHTML = True
   Else
     Me.btnHTMLUItslagen.Visible = False
     Me.btnKruisTabel.Visible = False
     Me.btnHTMLUItslagen.Enabled = False
     Me.btnKruisTabel.Enabled = False
     intUitvoerNaarHTML = False
End If


End Sub
