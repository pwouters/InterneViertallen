Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11521
    DatasheetFontHeight =11
    ItemSuffix =42
    Right =18735
    Bottom =12240
    DatasheetGridlinesColor =15132391
    Filter ="[ToernooiD]=1 and [id] = 3"
    RecSrcDt = Begin
        0x3ee22f0eb090e540
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="tblSessie"
    Caption ="Sessiegegevens"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnActivate ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =255
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
        Begin CheckBox
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
                    Name ="Bijschrift34"
                    Caption ="Sessiegegevens"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =3171
                    LayoutCachedHeight =1026
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2880
                    Left =5102
                    Top =283
                    Width =5151
                    Height =315
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="cboKiesSessie"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblSessie.id, tblSessie.Sessienaam, tblSessie.Sessienr FROM tblSessie WHE"
                        "RE [ToernooiD] = 1 ORDER BY tblSessie.id, tblSessie.Sessienr; "
                    ColumnWidths ="0;2835;567"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =5102
                    LayoutCachedTop =283
                    LayoutCachedWidth =10253
                    LayoutCachedHeight =598
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3401
                            Top =283
                            Width =1635
                            Height =320
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Kies Sessie_Etiket"
                            Caption ="Kies Sessie"
                            EventProcPrefix ="Kies_Sessie_Etiket"
                            GridlineColor =10921638
                            LayoutCachedLeft =3401
                            LayoutCachedTop =283
                            LayoutCachedWidth =5036
                            LayoutCachedHeight =603
                        End
                    End
                End
            End
        End
        Begin Section
            Height =6859
            Name ="Details"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2482
                    Height =315
                    ColumnWidth =1701
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="id"
                    ControlSource ="id"
                    GridlineColor =10921638

                    LayoutCachedLeft =2482
                    LayoutCachedWidth =4183
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =283
                            Width =2109
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="id_Bijschrift"
                            Caption ="id"
                            GridlineColor =10921638
                            LayoutCachedLeft =283
                            LayoutCachedWidth =2392
                            LayoutCachedHeight =330
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6847
                    Width =1530
                    Height =330
                    ColumnWidth =1530
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ToernooID"
                    ControlSource ="ToernooID"
                    GridlineColor =10921638

                    LayoutCachedLeft =6847
                    LayoutCachedWidth =8377
                    LayoutCachedHeight =330
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =4648
                            Width =2109
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ToernooID_Bijschrift"
                            Caption ="ToernooID"
                            GridlineColor =10921638
                            LayoutCachedLeft =4648
                            LayoutCachedWidth =6757
                            LayoutCachedHeight =330
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2552
                    Top =510
                    Width =3686
                    Height =600
                    ColumnWidth =3000
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Sessienaam"
                    ControlSource ="Sessienaam"
                    GridlineColor =10921638

                    LayoutCachedLeft =2552
                    LayoutCachedTop =510
                    LayoutCachedWidth =6238
                    LayoutCachedHeight =1110
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =345
                            Top =510
                            Width =2109
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Sessienaam_Bijschrift"
                            Caption ="Sessienaam"
                            GridlineColor =10921638
                            LayoutCachedLeft =345
                            LayoutCachedTop =510
                            LayoutCachedWidth =2454
                            LayoutCachedHeight =840
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2552
                    Top =1250
                    Width =1050
                    Height =330
                    ColumnWidth =1050
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Sessienr"
                    ControlSource ="Sessienr"
                    GridlineColor =10921638

                    LayoutCachedLeft =2552
                    LayoutCachedTop =1250
                    LayoutCachedWidth =3602
                    LayoutCachedHeight =1580
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =283
                            Top =1250
                            Width =2109
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Sessienr_Bijschrift"
                            Caption ="Sessienr"
                            GridlineColor =10921638
                            LayoutCachedLeft =283
                            LayoutCachedTop =1250
                            LayoutCachedWidth =2392
                            LayoutCachedHeight =1580
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9751
                    Top =1360
                    Width =1050
                    Height =330
                    ColumnWidth =1050
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Aantalspellen"
                    ControlSource ="Aantalspellen"
                    GridlineColor =10921638

                    LayoutCachedLeft =9751
                    LayoutCachedTop =1360
                    LayoutCachedWidth =10801
                    LayoutCachedHeight =1690
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6803
                            Top =1360
                            Width =2844
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Aantalspellen_Bijschrift"
                            Caption ="Aantalspellen"
                            GridlineColor =10921638
                            LayoutCachedLeft =6803
                            LayoutCachedTop =1360
                            LayoutCachedWidth =9647
                            LayoutCachedHeight =1690
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =2552
                    Top =2566
                    Width =3686
                    Height =600
                    ColumnWidth =3000
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Prefixkopjesscorestaat"
                    ControlSource ="Prefixkopjesscorestaat"
                    StatusBarText ="tekst links"
                    GridlineColor =10921638

                    LayoutCachedLeft =2552
                    LayoutCachedTop =2566
                    LayoutCachedWidth =6238
                    LayoutCachedHeight =3166
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =347
                            Top =2566
                            Width =2205
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Prefixkopjesscorestaat_Bijschrift"
                            Caption ="Scorestaat links"
                            GridlineColor =10921638
                            LayoutCachedLeft =347
                            LayoutCachedTop =2566
                            LayoutCachedWidth =2552
                            LayoutCachedHeight =2881
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2552
                    Top =3230
                    Width =3686
                    Height =600
                    ColumnWidth =3000
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="PrefixKopjeuitslagen"
                    ControlSource ="PrefixKopjeuitslagen"
                    StatusBarText ="tekst links"
                    GridlineColor =10921638

                    LayoutCachedLeft =2552
                    LayoutCachedTop =3230
                    LayoutCachedWidth =6238
                    LayoutCachedHeight =3830
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =345
                            Top =3230
                            Width =2115
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="PrefixKopjeuitslagen_Bijschrift"
                            Caption ="Uitslagen links"
                            GridlineColor =10921638
                            LayoutCachedLeft =345
                            LayoutCachedTop =3230
                            LayoutCachedWidth =2460
                            LayoutCachedHeight =3545
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2552
                    Top =3910
                    Width =3686
                    Height =600
                    ColumnWidth =3000
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Suffixkopjesscorestaat"
                    ControlSource ="Suffixkopjesscorestaat"
                    StatusBarText ="tekst rechts"
                    GridlineColor =10921638

                    LayoutCachedLeft =2552
                    LayoutCachedTop =3910
                    LayoutCachedWidth =6238
                    LayoutCachedHeight =4510
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =401
                            Top =3910
                            Width =2109
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Suffixkopjesscorestaat_Bijschrift"
                            Caption ="Scorestaat rechts"
                            GridlineColor =10921638
                            LayoutCachedLeft =401
                            LayoutCachedTop =3910
                            LayoutCachedWidth =2510
                            LayoutCachedHeight =4240
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2552
                    Top =4618
                    Width =3686
                    Height =600
                    ColumnWidth =3000
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="SuffixKopjeuitslagen"
                    ControlSource ="SuffixKopjeuitslagen"
                    StatusBarText ="tekst rechts"
                    GridlineColor =10921638

                    LayoutCachedLeft =2552
                    LayoutCachedTop =4618
                    LayoutCachedWidth =6238
                    LayoutCachedHeight =5218
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =347
                            Top =4618
                            Width =2109
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="SuffixKopjeuitslagen_Bijschrift"
                            Caption ="Uitslagen rechts"
                            GridlineColor =10921638
                            LayoutCachedLeft =347
                            LayoutCachedTop =4618
                            LayoutCachedWidth =2456
                            LayoutCachedHeight =4948
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2552
                    Top =5302
                    Width =3686
                    Height =600
                    ColumnWidth =3000
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Voettekst"
                    ControlSource ="Voettekst"
                    GridlineColor =10921638

                    LayoutCachedLeft =2552
                    LayoutCachedTop =5302
                    LayoutCachedWidth =6238
                    LayoutCachedHeight =5902
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =347
                            Top =5302
                            Width =2109
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Voettekst_Bijschrift"
                            Caption ="Voettekst"
                            GridlineColor =10921638
                            LayoutCachedLeft =347
                            LayoutCachedTop =5302
                            LayoutCachedWidth =2456
                            LayoutCachedHeight =5632
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2552
                    Top =5986
                    Width =3686
                    Height =600
                    ColumnWidth =3000
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Voetlink"
                    ControlSource ="Voetlink"
                    GridlineColor =10921638

                    LayoutCachedLeft =2552
                    LayoutCachedTop =5986
                    LayoutCachedWidth =6238
                    LayoutCachedHeight =6586
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =347
                            Top =5986
                            Width =2109
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Voetlink_Bijschrift"
                            Caption ="Voetlink"
                            GridlineColor =10921638
                            LayoutCachedLeft =347
                            LayoutCachedTop =5986
                            LayoutCachedWidth =2456
                            LayoutCachedHeight =6316
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9677
                    Top =3061
                    Width =1320
                    Height =330
                    ColumnWidth =1050
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ActivityID"
                    ControlSource ="ActivityID"
                    GridlineColor =10921638

                    LayoutCachedLeft =9677
                    LayoutCachedTop =3061
                    LayoutCachedWidth =10997
                    LayoutCachedHeight =3391
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6803
                            Top =3061
                            Width =2784
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ActivityID_Bijschrift"
                            Caption ="ActivityID (Step)"
                            GridlineColor =10921638
                            LayoutCachedLeft =6803
                            LayoutCachedTop =3061
                            LayoutCachedWidth =9587
                            LayoutCachedHeight =3391
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9737
                    Top =566
                    Width =1050
                    Height =330
                    ColumnWidth =1050
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="AantalTeams"
                    ControlSource ="AantalTeams"
                    GridlineColor =10921638

                    LayoutCachedLeft =9737
                    LayoutCachedTop =566
                    LayoutCachedWidth =10787
                    LayoutCachedHeight =896
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6803
                            Top =566
                            Width =2844
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="AantalTeams_Bijschrift"
                            Caption ="AantalTeams"
                            GridlineColor =10921638
                            LayoutCachedLeft =6803
                            LayoutCachedTop =566
                            LayoutCachedWidth =9647
                            LayoutCachedHeight =896
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =9737
                    Top =963
                    TabIndex =13
                    BorderColor =10921638
                    Name ="ByeTeam"
                    ControlSource ="ByeTeam"
                    StatusBarText ="Altijdhoogst genummerde team"
                    GridlineColor =10921638

                    LayoutCachedLeft =9737
                    LayoutCachedTop =963
                    LayoutCachedWidth =9997
                    LayoutCachedHeight =1203
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6803
                            Top =963
                            Width =2844
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ByeTeam_Bijschrift"
                            Caption ="ByeTeam"
                            GridlineColor =10921638
                            LayoutCachedLeft =6803
                            LayoutCachedTop =963
                            LayoutCachedWidth =9647
                            LayoutCachedHeight =1293
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9737
                    Top =1756
                    Width =1035
                    Height =330
                    ColumnWidth =1050
                    TabIndex =14
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="AantalWedstrijdenPerSessie"
                    ControlSource ="AantalWedstrijdenPerSessie"
                    StatusBarText ="Indien je indeelt op basis van de vorige ronde of random"
                    GridlineColor =10921638

                    LayoutCachedLeft =9737
                    LayoutCachedTop =1756
                    LayoutCachedWidth =10772
                    LayoutCachedHeight =2086
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6803
                            Top =1756
                            Width =2844
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="AantalWedstrijdenPerSessie_Bijschrift"
                            Caption ="Wedstrijden Per Sessie"
                            GridlineColor =10921638
                            LayoutCachedLeft =6803
                            LayoutCachedTop =1756
                            LayoutCachedWidth =9647
                            LayoutCachedHeight =2086
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    TextFontCharSet =177
                    Left =10204
                    Top =6236
                    Width =576
                    Height =576
                    TabIndex =15
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

                    LayoutCachedLeft =10204
                    LayoutCachedTop =6236
                    LayoutCachedWidth =10780
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
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =9700
                    Top =2607
                    Width =1821
                    Height =315
                    ColumnWidth =2235
                    TabIndex =16
                    BorderColor =10921638
                    ForeColor =3484194
                    Name ="Competitie"
                    ControlSource ="Competitie"
                    RowSourceType ="Value List"
                    RowSource ="1;\"Halve\";2:\"Hele"
                    ColumnWidths ="0;1701"
                    StatusBarText ="halve of hele"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =9700
                    LayoutCachedTop =2607
                    LayoutCachedWidth =11521
                    LayoutCachedHeight =2922
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6814
                            Top =2607
                            Width =2835
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift38"
                            Caption ="Competitie"
                            GridlineColor =10921638
                            LayoutCachedLeft =6814
                            LayoutCachedTop =2607
                            LayoutCachedWidth =9649
                            LayoutCachedHeight =2922
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =9700
                    Top =2210
                    Width =1821
                    Height =315
                    ColumnWidth =2490
                    TabIndex =17
                    BorderColor =10921638
                    ForeColor =3484194
                    Name ="wedstrijdvormID"
                    ControlSource ="wedstrijdvormID"
                    RowSourceType ="Value List"
                    RowSource ="0;\"Robin\";1;\"Zwitsers\";2;\"Deens\""
                    ColumnWidths ="0;1701"
                    StatusBarText ="robin, zwitsers, deens, eerste ronde 1-2, 3-4, tweede ronde 1-4, 3-2"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =9700
                    LayoutCachedTop =2210
                    LayoutCachedWidth =11521
                    LayoutCachedHeight =2525
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6814
                            Top =2210
                            Width =2835
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift39"
                            Caption ="wedstrijd Soort"
                            GridlineColor =10921638
                            LayoutCachedLeft =6814
                            LayoutCachedTop =2210
                            LayoutCachedWidth =9649
                            LayoutCachedHeight =2525
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =284
                    Top =1925
                    Width =2205
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Bijschrift41"
                    Caption ="Uitvoer HTML kopjes"
                    GridlineColor =10921638
                    LayoutCachedLeft =284
                    LayoutCachedTop =1925
                    LayoutCachedWidth =2489
                    LayoutCachedHeight =2240
                End
                Begin CommandButton
                    OverlapFlags =223
                    TextFontCharSet =177
                    Left =9637
                    Top =6236
                    Width =576
                    Height =576
                    TabIndex =18
                    ForeColor =-2147483630
                    Name ="btnNieuw"
                    Caption ="Knop83"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Record toevoegen"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b09880ff201010ff201010ff201010ff201010ff201010ff ,
                        0x201010ff201010ff201010ff201010ff201010ff201010ff201010ff00000000 ,
                        0x0000000000000000c0a090fffff8f0fffff8f0fffff0f0fffff0e0fff0e8e0ff ,
                        0xf0e8d0fff0e0d0fff0e0d0fff0e0d0fff0d8d0fff0d8d0ff201810ff00000000 ,
                        0x0000000000000000c0a090ffffffffffd07850ffd07840ffd07040ffc07040ff ,
                        0xc06840ffc06840ffc06840ffc07040ffa06040fff0e0d0ff403830ff00000000 ,
                        0x0000000000000000c0a890ffffffffffd07850fff0b8a0fff0b090fff0a880ff ,
                        0xf0a080fff09870fff09870fff0a880ffc09880fffff0f0ff909090ff00000000 ,
                        0x0000000000000000c0a890ffffffffffd07850ffd07850ffd07840ffd07040ff ,
                        0xc07040ffc07050ffd09070ff70b8c0ff90d8f0ff90f0ffff40c0e0ffa0f0ffff ,
                        0xa0e8ffff90d8f0ffc0a8a0fffffffffffffffffffffffffffffffffffff8f0ff ,
                        0xfff8f0fffff8f0fffff8f0ffb0e8ffff30b8e0ff80e8ffff60c8e0ff90f0ffff ,
                        0x30b8e0ffa0e8ffffc0a8a0ffc0a8a0ffc0a890ffc0a090ffc0a090ffc0a090ff ,
                        0xc09880ffc0a090ffd0c0b0ffa0e8ffff90f0ffffc0f8ffffb0e8f0ffc0f8ffff ,
                        0x90f0ffffa0f0ffff000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000020a8e0ff50c0e0ffb0e8f0fff0ffffffb0e8f0ff ,
                        0x50c0e0ff30b8e0ff000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000080e8ffc090f0ffffc0f8ffffb0e8f0ffc0f8ffff ,
                        0x90f0ffff90d8e0ff000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000050d8ff8030b8e0ff90f0ffff60c0e0ff90f0ffff ,
                        0x30b8e0ff50d0f080000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000030b0e0a040c8f09080e8ffc020b0e0ff70e8ffc0 ,
                        0x50d8f08030b0e080000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =9637
                    LayoutCachedTop =6236
                    LayoutCachedWidth =10213
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
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =223
                    TextFontCharSet =177
                    Left =9070
                    Top =6236
                    Width =576
                    Height =576
                    TabIndex =19
                    ForeColor =4210752
                    Name ="btnOpslaan"
                    Caption ="Knop146"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Record opslaan"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b09880ff201010ff201010ff201010ff201010ff201010ff ,
                        0x201010ff201010ff201810ff201810ff201810ff201810ff201810ff00000000 ,
                        0x0000000000000000c0a090fffff8f0fffff8f0fffff0f0fffff0e0fff0e8e0ff ,
                        0xf0e8d0fff0e0d0fff0e0d0fff0e0d0fff0e0d0fff0e0d0ff403830ff00000000 ,
                        0x0000000000000000c0a090ffffffffffd07850ffd07840ffd07040ffc07040ff ,
                        0xc07040ffc07850ffd09070ffd0a890ffd0a890fff0f0f0ff909090ff00000000 ,
                        0x0000000000000000c0a890ffffffffffd07850fff0b8a0fff0b090fff0a880ff ,
                        0xf0a880fff0b090ffe0b0a0ff804040ff703840ff703840ff703840ff703840ff ,
                        0x703840ff703840ffc0a890ffffffffffd07850ffd07850ffd07840ffd07040ff ,
                        0xd08050ffe0a890ffa05850ffc07870ff604840ffd0d8d0ffd0d8d0ff605040ff ,
                        0xc06060ff703840ffc0a8a0fffffffffffffffffffffffffffffffffffff8f0ff ,
                        0xfff8f0fffff8f0ffb06060ffe09090ff605040ff605040ff605040ff605040ff ,
                        0xc07070ff703840ffc0a8a0ffc0a8a0ffc0a890ffc0a090ffc0a090ffc0a090ff ,
                        0xc0a8a0ffe0d0c0ffc07070fff0a8b0ffe0a0a0ffe098a0ffe09090ffe08890ff ,
                        0xd08080ff703840ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000d08080ffd07070ffd06860ffd06860ffc05850ffc05850ff ,
                        0xb05040ff804040ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000d08890ffe07070ffffffffffffffffffffffffffffffffff ,
                        0xc05850ff904850ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000d09090ffe07070ffffffffffffffffffffffffffffffffff ,
                        0xd06860ffa05860ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000e0a0a0ffd09090ffd08890ffd08080ffc07070ffc06870ff ,
                        0xc06870ffc06860ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =9070
                    LayoutCachedTop =6236
                    LayoutCachedWidth =9646
                    LayoutCachedHeight =6812
                    Gradient =0
                    BackColor =-2147483612
                    BackThemeColorIndex =-1
                    BackTint =100.0
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
                    Enabled = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =177
                    Left =7937
                    Top =6236
                    Width =576
                    Height =576
                    TabIndex =20
                    ForeColor =4210752
                    Name ="btnUndo"
                    Caption ="Knop187"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Record ongedaan maken"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a0fb17d4aedb17d4ac0b17d4a7bb17d4a0c00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b17d4a0fb17d4affb17d4affb17d4affb17d4ae7b17d4a48 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000b17d4a0fb17d4a3fb17d4aa8b17d4affb17d4af9 ,
                        0xb17d4a3000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a96b17d4aff ,
                        0xb17d4ab100000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000b17d4a12b17d4aff ,
                        0xb17d4af000000000000000000000000000000000b17d4a5ab17d4afcb17d4aff ,
                        0xb17d4af9b17d4a4500000000000000000000000000000000b17d4a12b17d4aff ,
                        0xb17d4af6000000000000000000000000b17d4a42b17d4af9b17d4affb17d4afc ,
                        0xb17d4a510000000000000000000000000000000000000000b17d4a96b17d4aff ,
                        0xb17d4ac30000000000000000b17d4a36b17d4af6b17d4affb17d4affb17d4a5d ,
                        0x000000000000000000000000b17d4a12b17d4a42b17d4aa8b17d4affb17d4aff ,
                        0xb17d4a4b00000000b17d4a27b17d4aeab17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4affb17d4affb17d4affb17d4af9b17d4a78 ,
                        0x0000000000000000b17d4a24b17d4aeab17d4affb17d4affb17d4affb17d4aff ,
                        0xb17d4affb17d4affb17d4affb17d4aedb17d4accb17d4a90b17d4a2400000000 ,
                        0x000000000000000000000000b17d4a2db17d4aedb17d4affb17d4affb17d4a5a ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000b17d4a36b17d4af3b17d4affb17d4af9 ,
                        0xb17d4a3c00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000b17d4a42b17d4af6b17d4aff ,
                        0xb17d4aeab17d4a24000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =7937
                    LayoutCachedTop =6236
                    LayoutCachedWidth =8513
                    LayoutCachedHeight =6812
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
                    OverlapFlags =215
                    TextFontCharSet =177
                    Left =8503
                    Top =6236
                    Width =576
                    Height =576
                    TabIndex =21
                    ForeColor =4210752
                    Name ="btnDelete"
                    Caption ="Knop188"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Record verwijderen"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000008080803980808096868686d6 ,
                        0x828282f7808080ff828282f8868686d68585859c808080390000000000000000 ,
                        0x0000000000000000000000000000000000000000808080ff808080ff808080ff ,
                        0x808080ff808080ff808080ff808080ff808080ff808080ff0000000000000000 ,
                        0x0000000000000000000000000000000080808006818181fbcbcbcbffe6e6e6ff ,
                        0xf9f9f9fffffffffff9f9f9ffe6e6e6ffcbcbcbff808080ff0000000000000000 ,
                        0x000000000000000000000000000000008080801e898989edffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff868686fc8080801500000000 ,
                        0x0000000000000000000000000000000080808036939393e6ffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff929292fc8080802d00000000 ,
                        0x00000000000000000000000000000000808080519f9f9fe3ffffffffe0e5d4ff ,
                        0x90a468ff90a468ff90a468ffdce2cfffffffffffa0a0a0fe8383834400000000 ,
                        0x0000000000000000000000000000000080808069afafafe5ffffffffadbc8fff ,
                        0xcfd7bdffffffffffd8dfcaffa6b686fffefefdffadadadff9393936a00000000 ,
                        0x0000000000000000000000000000000080808081bebebeecfffffffff6f8f3ff ,
                        0xa0b17dffc2cdacff9eb07cfff3f5efffffffffffb9b9b9ff9999998e00000000 ,
                        0x000000000000000000000000000000008080809ccececef9ffffffffffffffff ,
                        0xeff2eaffbbc7a3ffebefe4ffffffffffffffffffc5c5c5ff979797ad00000000 ,
                        0x00000000000000000000000000000000838383b8ddddddffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffcfcfcfff969696c700000000 ,
                        0x00000000000000000000000000000000848484d3e9e9e9ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffddddddff929292dc00000000 ,
                        0x00000000000000000000000000000000838383edccccccff9b9b9bff808080ff ,
                        0x808080ff808080ff808080ff808080ff9b9b9bffc2c2c2ff8d8d8ded00000000 ,
                        0x00000000000000000000000080808003808080ffa4a4a4ffdededeffffffffff ,
                        0xffffffffffffffffffffffffffffffffdededeffa4a4a4ff808080ff00000000 ,
                        0x00000000000000000000000000000000808080ff9c9c9ccddbdbdbe7ffffffff ,
                        0xffffffffffffffffffffffffffffffffddddddf4a4a4a4dd808080ff00000000 ,
                        0x00000000000000000000000000000000000000008080804e808080bd808080ff ,
                        0x808080ff808080ff808080ff808080ff808080bd8080804e0000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =8503
                    LayoutCachedTop =6236
                    LayoutCachedWidth =9079
                    LayoutCachedHeight =6812
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
Dim MyKey, MyKeyIs As String

Private Sub btnSluiten_Click()
DoCmd.Close
End Sub

Private Sub cboKiesSessie_AfterUpdate()
 Dim rs As Recordset
    Dim X
   Dim criterium As String
   If Me.Dirty Then
       X = fnSaveRecords
    End If
    Set rs = Recordset
   
    criterium = MyKeyIs & cboKiesSessie
    rs.FindFirst criterium
    
    If rs.NoMatch Then
    MsgBox "Geen Sessie bekend in de database"
    cboKiesSessie = ""
    Else
    Me.Bookmark = rs.Bookmark
    cboKiesSessie = ""
    End If
End Sub


Private Sub btnDelete_Click()

Dim rs As Recordset

If Me.NewRecord Then
 Call MeControlEnabled(True)
 Call GotoLastCurrentRecord(Me.Form.name, MyKey, lngPK)
 'DoCmd.GoToRecord , , acGoTo, recordteller
Else
    Set rs = Recordset
    With rs
      .FindFirst MyKeyIs & lngPK   ' Use actual field names
      .Delete
      If Not .BOF Then
          .MovePrevious
          Me.Bookmark = .Bookmark
      End If
    End With
 End If
End Sub





Private Sub btnNieuw_Click()



 If Me.Dirty Then
       X = fnSaveRecords
 End If
    DoCmd.RunCommand acCmdRecordsGoToNew
    Me.ToernooID = lngToernooi
    Me.Sessienr = DMax("Sessienr", "tblSessie", "[ToernooiD] = " & lngToernooi) + 1
    
 Call MeControlEnabled(True)
 Me.cboKiesSessie.Enabled = False

End Sub

Private Sub btnOpslaan_Click()
 Dim X
 If Me.Dirty Then
       X = fnSaveRecords
 End If
  Me.cboKiesSessie.SetFocus
    Me.btnOpslaan.Visible = False
    Me.btnUndo.Visible = False
    Me.btnOpslaan.Enabled = False
    Me.btnUndo.Enabled = False
     lngSessie = Me.Id
 'Call MeControlEnabled(True)
 Me.cboKiesSessie.Enabled = True
End Sub

Private Sub btnUndo_Click()
On Error Resume Next
If Me.NewRecord Then
     DoCmd.RunCommand acCmdUndo
    ' Call MeControlEnabled(True)
     Me.cboKiesSessie.Enabled = True
     Call GotoLastCurrentRecord(Me.Form.name, MyKey, lngPK)
     'DoCmd.GoToRecord , , acGoTo, recordteller
    Else
     DoCmd.RunCommand acCmdUndo
End If
    Me.cboKiesSessie.SetFocus
    Me.btnOpslaan.Visible = False
    Me.btnUndo.Visible = False
    Me.btnOpslaan.Enabled = False
    Me.btnUndo.Enabled = False
     
End Sub




Private Sub Form_Current()
    Me.btnUndo.Enabled = False
    Me.btnUndo.Visible = False
    
  If Me.NewRecord = False Then
      lngPK = Me.Id
      lngSessie = Me.Id
      lngToernooi = Me.ToernooID
  End If
End Sub

Private Sub Form_Dirty(Cancel As Integer)
    Me.btnOpslaan.Visible = True
    Me.btnUndo.Visible = True
    Me.btnOpslaan.Enabled = True
    Me.btnUndo.Enabled = True
End Sub




Private Sub Form_Activate()
Dim sql As String

MyKey = "id"
MyKeyIs = MyKey & " = "

If CurrentProject.AllForms("Start_VT").IsLoaded = True Then
    Me.Filter = "[ToernooiD]=" & lngToernooi & " and [id] = " & lngSessie
    Me.FilterOn = True
    sql = "SELECT tblSessie.id, tblSessie.Sessienaam, tblSessie.Sessienr From tblSessie Where [ToernooiD] = " & lngToernooi & " ORDER BY tblSessie.id, tblSessie.Sessienr;"
    Me.cboKiesSessie.RowSource = sql
    Me.cboKiesSessie.Requery
Else
    If CurrentProject.AllForms("Toernooigegevens").IsLoaded = True Then
        Me.Filter = "[ToernooiD]=" & lngToernooi
        Me.FilterOn = True
        sql = "SELECT tblSessie.id, tblSessie.Sessienaam, tblSessie.Sessienr From tblSessie Where [ToernooiD] = " & lngToernooi & " ORDER BY tblSessie.id, tblSessie.Sessienr;"
        Me.cboKiesSessie.RowSource = sql
        Me.cboKiesSessie.Requery
    End If
End If


End Sub

Private Sub Form_Open(Cancel As Integer)
MyKey = "id"
MyKeyIs = MyKey & " = "
Dim criterium As String

Dim sql As String
Dim rs As Recordset
sql = "SELECT tblSessie.id, tblSessie.Sessienaam, tblSessie.Sessienr From tblSessie "

If CurrentProject.AllForms("Start_VT").IsLoaded = False Then
        Me.cboKiesSessie.Visible = True
        Me.cboKiesSessie.Enabled = True
    If CurrentProject.AllForms("Toernooigegevens").IsLoaded = False Then
        'je opent alleen de sessieform
        lngToernooi = 0
        sql = "SELECT tblSessie.id, tblSessie.Sessienaam, tblSessie.Sessienr From tblSessie ORDER BY tblSessie.id, tblSessie.Sessienr;"
        Me.cboKiesSessie.RowSource = sql
        Me.cboKiesSessie.Requery

        Me.Filter = ""
        Me.FilterOn = False
    End If
 End If


End Sub
Sub MeControlEnabled(JaNee As Integer)
' Me.btnVolgende.Enabled = JaNee
 Me.btnNieuw.Enabled = JaNee
 'Me.btnVorige.Enabled = JaNee
 'Me.btnEerste.Enabled = JaNee
' Me.btnLaatste.Enabled = JaNee
 'Me.btnDelete.Enabled = JaNee
End Sub

Sub MeControlVisible(JaNee As Integer)
' Me.btnVolgende.Visible = JaNee
 Me.btnNieuw.Visible = JaNee
' Me.btnVorige.Visible = JaNee
'Me.btnEerste.Visible = JaNee
' Me.btnLaatste.Visible = JaNee
 'Me.btnDelete.Enabled = JaNee
End Sub
