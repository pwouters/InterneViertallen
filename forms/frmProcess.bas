Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10714
    DatasheetFontHeight =11
    ItemSuffix =28
    Right =18735
    Bottom =12240
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
            Height =7889
            Name ="Details"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =915
                    Top =345
                    Width =5775
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Bijschrift0"
                    Caption ="Process  Viertallen"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =915
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
                            Left =907
                            Top =964
                            Width =2310
                            Height =320
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Kies Toernooi_Etiket"
                            Caption ="Kies Toernooi"
                            EventProcPrefix ="Kies_Toernooi_Etiket"
                            GridlineColor =10921638
                            LayoutCachedLeft =907
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
                    Left =8561
                    Top =283
                    Height =315
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ID"
                    ControlSource ="ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =8561
                    LayoutCachedTop =283
                    LayoutCachedWidth =10262
                    LayoutCachedHeight =598
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =6860
                            Top =283
                            Width =270
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift4"
                            Caption ="ID"
                            GridlineColor =10921638
                            LayoutCachedLeft =6860
                            LayoutCachedTop =283
                            LayoutCachedWidth =7130
                            LayoutCachedHeight =598
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =911
                    Top =2267
                    Width =2496
                    Height =568
                    TabIndex =2
                    ForeColor =4210752
                    Name ="btnOphalenScores"
                    Caption ="Process Scorestaten"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =911
                    LayoutCachedTop =2267
                    LayoutCachedWidth =3407
                    LayoutCachedHeight =2835
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
                    Left =911
                    Top =2834
                    Width =2496
                    Height =568
                    TabIndex =3
                    ForeColor =4210752
                    Name ="btnHTMLUItslagen"
                    Caption ="Process HTML Uitslagen"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =911
                    LayoutCachedTop =2834
                    LayoutCachedWidth =3407
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
                    OverlapFlags =87
                    Left =911
                    Top =3401
                    Width =2481
                    Height =568
                    TabIndex =4
                    ForeColor =4210752
                    Name ="btnKruisTabel"
                    Caption ="Process HTML Kruistabel"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =911
                    LayoutCachedTop =3401
                    LayoutCachedWidth =3392
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
                Begin OptionButton
                    OverlapFlags =85
                    Left =3628
                    Top =2411
                    TabIndex =5
                    BorderColor =10921638
                    Name ="optAlle"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3628
                    LayoutCachedTop =2411
                    LayoutCachedWidth =3888
                    LayoutCachedHeight =2651
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =3968
                            Top =2381
                            Width =555
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblAlle"
                            Caption ="alle"
                            GridlineColor =10921638
                            LayoutCachedLeft =3968
                            LayoutCachedTop =2381
                            LayoutCachedWidth =4523
                            LayoutCachedHeight =2696
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6690
                    Top =2381
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

                    LayoutCachedLeft =6690
                    LayoutCachedTop =2381
                    LayoutCachedWidth =9576
                    LayoutCachedHeight =2696
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5544
                            Top =2381
                            Width =975
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblKiesTeam"
                            Caption ="Team"
                            GridlineColor =10921638
                            LayoutCachedLeft =5544
                            LayoutCachedTop =2381
                            LayoutCachedWidth =6519
                            LayoutCachedHeight =2696
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
                            Left =907
                            Top =1360
                            Width =2325
                            Height =320
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Sessienr_Etiket"
                            Caption ="Sessienr"
                            GridlineColor =10921638
                            LayoutCachedLeft =907
                            LayoutCachedTop =1360
                            LayoutCachedWidth =3232
                            LayoutCachedHeight =1680
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =911
                    Top =4535
                    Width =2496
                    Height =568
                    TabIndex =8
                    ForeColor =4210752
                    Name ="btnImportTeams"
                    Caption ="Import Teams"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =911
                    LayoutCachedTop =4535
                    LayoutCachedWidth =3407
                    LayoutCachedHeight =5103
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
                    Left =911
                    Top =5102
                    Width =2496
                    Height =568
                    TabIndex =9
                    ForeColor =4210752
                    Name ="btnImportOpstelling"
                    Caption ="Import Opstelling"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =911
                    LayoutCachedTop =5102
                    LayoutCachedWidth =3407
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
                    OverlapFlags =87
                    Left =911
                    Top =5669
                    Width =2496
                    Height =568
                    TabIndex =10
                    ForeColor =4210752
                    Name ="btnImportIndeling"
                    Caption ="Import Indeling"
                    GridlineColor =10921638

                    LayoutCachedLeft =911
                    LayoutCachedTop =5669
                    LayoutCachedWidth =3407
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
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =911
                    Top =4138
                    Width =2490
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblImport"
                    Caption ="Import vanuit werkbestand"
                    GridlineColor =10921638
                    LayoutCachedLeft =911
                    LayoutCachedTop =4138
                    LayoutCachedWidth =3401
                    LayoutCachedHeight =4453
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =3968
                    Top =4535
                    Width =2556
                    Height =568
                    TabIndex =11
                    ForeColor =4210752
                    Name ="btnExportTeams"
                    Caption ="Export Teams"
                    GridlineColor =10921638

                    LayoutCachedLeft =3968
                    LayoutCachedTop =4535
                    LayoutCachedWidth =6524
                    LayoutCachedHeight =5103
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
                    Left =3968
                    Top =5102
                    Width =2556
                    Height =568
                    TabIndex =12
                    ForeColor =4210752
                    Name ="btnExportOpstelling"
                    Caption ="Export Opstelling"
                    GridlineColor =10921638

                    LayoutCachedLeft =3968
                    LayoutCachedTop =5102
                    LayoutCachedWidth =6524
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
                    OverlapFlags =87
                    Left =3968
                    Top =5669
                    Width =2556
                    Height =568
                    TabIndex =13
                    ForeColor =4210752
                    Name ="btnExportIndeling"
                    Caption ="Export Indeling"
                    GridlineColor =10921638

                    LayoutCachedLeft =3968
                    LayoutCachedTop =5669
                    LayoutCachedWidth =6524
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
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =3968
                    Top =4138
                    Width =2550
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Bijschrift25"
                    Caption ="Export naar werkbestand"
                    GridlineColor =10921638
                    LayoutCachedLeft =3968
                    LayoutCachedTop =4138
                    LayoutCachedWidth =6518
                    LayoutCachedHeight =4453
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =177
                    Left =9070
                    Top =6236
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

                    LayoutCachedLeft =9070
                    LayoutCachedTop =6236
                    LayoutCachedWidth =9646
                    LayoutCachedHeight =6812
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
                    Left =7370
                    Top =4535
                    Width =2271
                    Height =568
                    TabIndex =15
                    ForeColor =4210752
                    Name ="btnNewToernooi"
                    Caption ="Nieuw Toernooi"
                    GridlineColor =10921638

                    LayoutCachedLeft =7370
                    LayoutCachedTop =4535
                    LayoutCachedWidth =9641
                    LayoutCachedHeight =5103
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
                    Left =7370
                    Top =5102
                    Width =2271
                    Height =568
                    TabIndex =16
                    ForeColor =4210752
                    Name ="btnNieuweSessie"
                    Caption ="Nieuwe Sessie"
                    GridlineColor =10921638

                    LayoutCachedLeft =7370
                    LayoutCachedTop =5102
                    LayoutCachedWidth =9641
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
    lngSessie = DLookup("id", "tblSessie", "Sessienr=" & Me.cboKiesSessie & " and ToernooiD = " & lngToernooi)
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


If Me.optAlle = True Then
     Me.lblKiesTeam.Visible = False
     Me.cboKiesTeam.Visible = False
   Else
     Me.lblKiesTeam.Visible = True
     Me.cboKiesTeam.Visible = True
End If

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
