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
    Width =14229
    DatasheetFontHeight =11
    ItemSuffix =48
    Right =13995
    Bottom =10470
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0x9b1f09e7d98fe540
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="tblToernooi"
    Caption ="Toernooigegevens"
    OnCurrent ="[Event Procedure]"
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
                    Width =3690
                    Height =969
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Bijschrift32"
                    Caption ="Toernooigegevens"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =3747
                    LayoutCachedHeight =1026
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2865
                    Left =6236
                    Top =226
                    Width =4656
                    Height =315
                    ColumnOrder =0
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="cboKiesToernooi"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [tblToernooi].[ID], [tblToernooi].[ToernooiNaam] FROM tblToernooi; "
                    ColumnWidths ="0;2865"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6236
                    LayoutCachedTop =226
                    LayoutCachedWidth =10892
                    LayoutCachedHeight =541
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3965
                            Top =226
                            Width =1995
                            Height =320
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Kies Toernooi_Etiket"
                            Caption ="Kies Toernooi"
                            EventProcPrefix ="Kies_Toernooi_Etiket"
                            GridlineColor =10921638
                            LayoutCachedLeft =3965
                            LayoutCachedTop =226
                            LayoutCachedWidth =5960
                            LayoutCachedHeight =546
                        End
                    End
                End
            End
        End
        Begin Section
            Height =6746
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
                    Left =1857
                    Top =342
                    Height =315
                    ColumnWidth =1701
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ID"
                    ControlSource ="ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =1857
                    LayoutCachedTop =342
                    LayoutCachedWidth =3558
                    LayoutCachedHeight =657
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =57
                            Top =342
                            Width =1710
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ID_Bijschrift"
                            Caption ="ID"
                            GridlineColor =10921638
                            LayoutCachedLeft =57
                            LayoutCachedTop =342
                            LayoutCachedWidth =1767
                            LayoutCachedHeight =672
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1857
                    Top =741
                    Width =4710
                    Height =600
                    ColumnWidth =3000
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ToernooiNaam"
                    ControlSource ="ToernooiNaam"
                    GridlineColor =10921638

                    LayoutCachedLeft =1857
                    LayoutCachedTop =741
                    LayoutCachedWidth =6567
                    LayoutCachedHeight =1341
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =741
                            Width =1710
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ToernooiNaam_Bijschrift"
                            Caption ="ToernooiNaam"
                            GridlineColor =10921638
                            LayoutCachedLeft =57
                            LayoutCachedTop =741
                            LayoutCachedWidth =1767
                            LayoutCachedHeight =1071
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1857
                    Top =1425
                    Width =4710
                    Height =1140
                    ColumnWidth =3000
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="WORKFOLDER"
                    ControlSource ="WORKFOLDER"
                    OnMouseDown ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1857
                    LayoutCachedTop =1425
                    LayoutCachedWidth =6567
                    LayoutCachedHeight =2565
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =1425
                            Width =1710
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="WORKFOLDER_Bijschrift"
                            Caption ="Werkfolder"
                            GridlineColor =10921638
                            LayoutCachedLeft =57
                            LayoutCachedTop =1425
                            LayoutCachedWidth =1767
                            LayoutCachedHeight =1755
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1812
                    Top =2679
                    Width =4755
                    Height =600
                    ColumnWidth =3000
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="WORKFILE"
                    ControlSource ="WORKFILE"
                    OnMouseDown ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1812
                    LayoutCachedTop =2679
                    LayoutCachedWidth =6567
                    LayoutCachedHeight =3279
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =57
                            Top =2679
                            Width =1635
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="WORKFILE_Bijschrift"
                            Caption ="Werkbestand"
                            GridlineColor =10921638
                            LayoutCachedLeft =57
                            LayoutCachedTop =2679
                            LayoutCachedWidth =1692
                            LayoutCachedHeight =2994
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8318
                    Top =737
                    Width =4710
                    Height =1140
                    ColumnWidth =3000
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="STEPDATA"
                    ControlSource ="STEPDATA"
                    GridlineColor =10921638

                    LayoutCachedLeft =8318
                    LayoutCachedTop =737
                    LayoutCachedWidth =13028
                    LayoutCachedHeight =1877
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6803
                            Top =737
                            Width =1425
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="STEPDATA_Bijschrift"
                            Caption ="Step admin"
                            GridlineColor =10921638
                            LayoutCachedLeft =6803
                            LayoutCachedTop =737
                            LayoutCachedWidth =8228
                            LayoutCachedHeight =1067
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8318
                    Top =1991
                    Width =4710
                    Height =1140
                    ColumnWidth =3000
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="STEPRESULTS"
                    ControlSource ="STEPRESULTS"
                    GridlineColor =10921638

                    LayoutCachedLeft =8318
                    LayoutCachedTop =1991
                    LayoutCachedWidth =13028
                    LayoutCachedHeight =3131
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6803
                            Top =1991
                            Width =1425
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="STEPRESULTS_Bijschrift"
                            Caption ="Step Results"
                            GridlineColor =10921638
                            LayoutCachedLeft =6803
                            LayoutCachedTop =1991
                            LayoutCachedWidth =8228
                            LayoutCachedHeight =2321
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8318
                    Top =3458
                    Width =4710
                    Height =600
                    ColumnWidth =3000
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="LOCALSITE"
                    ControlSource ="LOCALSITE"
                    GridlineColor =10921638

                    LayoutCachedLeft =8318
                    LayoutCachedTop =3458
                    LayoutCachedWidth =13028
                    LayoutCachedHeight =4058
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6803
                            Top =3458
                            Width =1425
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="LOCALSITE_Bijschrift"
                            Caption ="Lokale site"
                            GridlineColor =10921638
                            LayoutCachedLeft =6803
                            LayoutCachedTop =3458
                            LayoutCachedWidth =8228
                            LayoutCachedHeight =3788
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8318
                    Top =4142
                    Width =4710
                    Height =600
                    ColumnWidth =3000
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="LOCALHTML"
                    ControlSource ="LOCALHTML"
                    OnMouseDown ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =8318
                    LayoutCachedTop =4142
                    LayoutCachedWidth =13028
                    LayoutCachedHeight =4742
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6803
                            Top =4142
                            Width =1425
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="LOCALHTML_Bijschrift"
                            Caption ="Opslag HTML"
                            GridlineColor =10921638
                            LayoutCachedLeft =6803
                            LayoutCachedTop =4142
                            LayoutCachedWidth =8228
                            LayoutCachedHeight =4472
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1814
                    Top =3514
                    Width =4776
                    Height =600
                    ColumnWidth =3000
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="WORKTEMPLATE"
                    ControlSource ="WORKTEMPLATE"
                    StatusBarText ="excelblad om mee te starten"
                    OnMouseDown ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1814
                    LayoutCachedTop =3514
                    LayoutCachedWidth =6590
                    LayoutCachedHeight =4114
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =111
                            Top =3515
                            Width =1582
                            Height =317
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="WORKTEMPLATE_Bijschrift"
                            Caption ="Modelbestand"
                            GridlineColor =10921638
                            LayoutCachedLeft =111
                            LayoutCachedTop =3515
                            LayoutCachedWidth =1693
                            LayoutCachedHeight =3832
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =566
                    Top =5102
                    Width =2265
                    Height =570
                    TabIndex =9
                    ForeColor =4210752
                    Name ="btnSessiegegevens"
                    Caption ="Sessie gegevens"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =5102
                    LayoutCachedWidth =2831
                    LayoutCachedHeight =5672
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
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =177
                    Left =10204
                    Top =5669
                    Width =576
                    Height =576
                    TabIndex =10
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

                    LayoutCachedLeft =10204
                    LayoutCachedTop =5669
                    LayoutCachedWidth =10780
                    LayoutCachedHeight =6245
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
                    OverlapFlags =215
                    TextFontCharSet =177
                    Left =10771
                    Top =5669
                    Width =576
                    Height =576
                    TabIndex =11
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

                    LayoutCachedLeft =10771
                    LayoutCachedTop =5669
                    LayoutCachedWidth =11347
                    LayoutCachedHeight =6245
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
                    Left =9636
                    Top =5669
                    Width =576
                    Height =576
                    TabIndex =12
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

                    LayoutCachedLeft =9636
                    LayoutCachedTop =5669
                    LayoutCachedWidth =10212
                    LayoutCachedHeight =6245
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
                    OverlapFlags =93
                    TextFontCharSet =177
                    Left =8503
                    Top =5669
                    Width =576
                    Height =576
                    TabIndex =13
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

                    LayoutCachedLeft =8503
                    LayoutCachedTop =5669
                    LayoutCachedWidth =9079
                    LayoutCachedHeight =6245
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
                    Left =9069
                    Top =5669
                    Width =576
                    Height =576
                    TabIndex =14
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

                    LayoutCachedLeft =9069
                    LayoutCachedTop =5669
                    LayoutCachedWidth =9645
                    LayoutCachedHeight =6245
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
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8320
                    Top =4989
                    Width =3531
                    Height =315
                    TabIndex =15
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="PREFIX"
                    ControlSource ="PREFIX"
                    GridlineColor =10921638

                    LayoutCachedLeft =8320
                    LayoutCachedTop =4989
                    LayoutCachedWidth =11851
                    LayoutCachedHeight =5304
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6816
                            Top =4988
                            Width =1395
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift46"
                            Caption ="Prefix  files"
                            GridlineColor =10921638
                            LayoutCachedLeft =6816
                            LayoutCachedTop =4988
                            LayoutCachedWidth =8211
                            LayoutCachedHeight =5303
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1814
                    Top =4308
                    Height =315
                    TabIndex =16
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="AANTALSESSIES"
                    ControlSource ="AANTALSESSIES"
                    StatusBarText ="Aantal zittingen van tig rondes"
                    GridlineColor =10921638

                    LayoutCachedLeft =1814
                    LayoutCachedTop =4308
                    LayoutCachedWidth =3515
                    LayoutCachedHeight =4623
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =56
                            Top =4308
                            Width =1515
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift47"
                            Caption ="AANTALSESSIES"
                            GridlineColor =10921638
                            LayoutCachedLeft =56
                            LayoutCachedTop =4308
                            LayoutCachedWidth =1571
                            LayoutCachedHeight =4623
                        End
                    End
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
Dim MyKey, MyKeyIs As String

Private Sub btnSessiegegevens_Click()
    Dim db As Database
    Dim rs As Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("Select * from tblSessie where ToernooiD = " & Me.Id)
    If rs.BOF And rs.EOF Then
    
        ' voeg een nieuw sessie record
        rs.AddNew
        rs!ToernooID = Me.Id
        rs!Sessienr = 1
        rs!Sessienaam = Me.ToernooiNaam
        rs.Update
        rs.Close
        db.Close
    Else
        If lngSessie = 0 Then
            lngSessie = rs!Id
        End If
        rs.Close
        db.Close
    End If
    
    
    lngToernooi = Me.Id
    
    'test even of al een sessie is anders maak een nieuwe sessie aan
    
    DoCmd.OpenForm "Sessiegegevens", acNormal, , "[ToernooiD] =" & lngToernooi
    
End Sub

Private Sub btnSluiten_Click()
   Dim x
   If Me.Dirty Then
       x = fnSaveRecords
    End If
    If CurrentProject.AllForms("Start_VT").IsLoaded = False Then
        DoCmd.Close
    Else
        DoCmd.BrowseTo acBrowseToForm, "frmBegin"
   End If
End Sub

Private Sub cboKiesToernooi_AfterUpdate()
Dim rs As Recordset
    Dim x
   Dim Criterium As String
   If Me.Dirty Then
      x = fnSaveRecords
    End If
    Set rs = Recordset
   
    Criterium = MyKeyIs & cboKiesToernooi
    rs.FindFirst Criterium
    
    If rs.NoMatch Then
    MsgBox "Geen lijst bekend in de database"
    cboKiesToernooi = ""
    Else
    Me.Bookmark = rs.Bookmark
    cboKiesToernooi = ""
    End If
    lngToernooi = Me.Id
    lngSessie = DLookup("id", "tblSessie", "Sessienr=" & 1 & " and ToernooiD = " & lngToernooi)
    Call InitAll(lngToernooi, lngSessie)
End Sub




Private Sub Form_Open(Cancel As Integer)

Dim rs As Recordset
Dim Criterium

MyKey = "ID"
MyKeyIs = MyKey & " = "
Me.cboKiesToernooi = ""

If CurrentProject.AllForms("Start_VT").IsLoaded = False Then
    Me.cboKiesToernooi.Visible = True
    Me.cboKiesToernooi.Enabled = True
End If


If lngToernooi = 0 Then
        lngToernooiOld = 0
        lngSessieOld = 0
        lngToernooi = 1
        lngSessie = 1
        Call InitAll(lngToernooi, lngSessie)
        Me.cboKiesToernooi.Value = lngToernooi
     Else
      Me.cboKiesToernooi.Value = lngToernooi
      Call InitAll(lngToernooi, lngSessie)
End If


Set rs = Me.RecordsetClone
   
Criterium = MyKeyIs & cboKiesToernooi
rs.FindFirst Criterium

If rs.NoMatch Then
MsgBox "Geen lijst bekend in de database"
cboKiesToernooi = ""
Else
Me.Bookmark = rs.Bookmark
cboKiesToernooi = ""
End If
rs.Close

strExcel_Folder = Me.WORKFOLDER
strHTML_Folder = Me.LOCALHTML
strTemplate_File = Me.WORKTEMPLATE
Call MeControlEnabled(True)
Call MeControlVisible(True)
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
       x = fnSaveRecords
 End If
 
 DoCmd.RunCommand acCmdRecordsGoToNew
 Call MeControlEnabled(True)
 Me.cboKiesToernooi.Enabled = False

' saveas model

'Nieuwe bestandsnaam
'nieuwe excelfolder

'indien escape ga terug naar het oude record






End Sub

Private Sub btnOpslaan_Click()
 Dim x
 If Me.Dirty Then
       x = fnSaveRecords
 End If
 Me.cboKiesToernooi.Enabled = True
 Me.cboKiesToernooi.Visible = True
  Me.cboKiesToernooi.SetFocus
    Me.btnOpslaan.Visible = False
    Me.btnUndo.Visible = False
    Me.btnOpslaan.Enabled = False
    Me.btnUndo.Enabled = False
      
 
 'Call MeControlEnabled(True)
 Me.cboKiesToernooi.Enabled = True
End Sub

Private Sub btnUndo_Click()
On Error Resume Next
If Me.NewRecord Then
     DoCmd.RunCommand acCmdUndo
    ' Call MeControlEnabled(True)
     Me.cboKiesToernooi.Enabled = True
     Call GotoLastCurrentRecord(Me.Form.name, MyKey, lngPK)
     'DoCmd.GoToRecord , , acGoTo, recordteller
    Else
     DoCmd.RunCommand acCmdUndo
End If
    Me.cboKiesToernooi.SetFocus
    Me.btnOpslaan.Visible = False
    Me.btnUndo.Visible = False
    Me.btnOpslaan.Enabled = False
    Me.btnUndo.Enabled = False
     
End Sub




Private Sub Form_Current()
    Me.btnUndo.Enabled = False
    Me.btnUndo.Visible = False
    If Not IsNull(Me.Id) Then
        lngToernooi = Me.Id
    End If
    If Me.NewRecord = False Then
        lngToernooi = Me.Id
        lngPK = Me.Id
    End If
End Sub



'Dim Ctrl As Control

'MyKey = "clntID"
'MyKeyIs = MyKey & " = "

'For Each Ctrl In Me.Controls
'    If Ctrl.Name <> "cboKiesclient" And Ctrl.Name <> "lblKiesclient" And Ctrl.Name <> "lblKoptekst" And Not Ctrl.Name Like "Nav*" Then
'        Ctrl.Visible = False
 '   End If
    
'Next
'CtrlVisible = False
'Set Ctrl = Nothing



'End Sub

Private Sub Form_Dirty(Cancel As Integer)
    Me.btnOpslaan.Visible = True
    Me.btnUndo.Visible = True
    Me.btnOpslaan.Enabled = True
    Me.btnUndo.Enabled = True
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

Private Sub LOCALHTML_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
 'If right-button clicked
    If Button = 1 Then
            Dim strFolder As String
            strFolder = GetFolderName(strHTML_Folder, "html")
            If Not strFolder = "" Then
                Me.LOCALHTML = strFolder
            End If
    End If
End Sub

Private Sub WORKFILE_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
' If right-button clicked
    If Button = 1 Then
            Dim strFile As String
            strFile = GetFileName(strExcel_Folder)
            If Not strFile = "" Then
                'strip path
                strFile = FileNameFromPath(strFile)
                Me.WORKFILE = strFile
            End If
    End If
End Sub

Private Sub WORKFOLDER_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 1 Then
            Dim strFolder As String
            strFolder = GetFolderName(strExcel_Folder)
            If Not strFolder = "" Then
                Me.WORKFOLDER = strFolder
                strExcel_Folder = strFolder
            End If
    End If
End Sub

Private Sub WORKTEMPLATE_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
' If right-button clicked
    If Button = 1 Then
            Dim strFile As String
            strFile = GetFileName(strTemplate_File)
            If Not strFile = "" Then
                'strip path
                'strFile = FileNameFromPath(strFile)
                Me.WORKTEMPLATE = strFile
            End If
    End If
End Sub
