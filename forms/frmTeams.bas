Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10209
    DatasheetFontHeight =11
    ItemSuffix =25
    Left =2520
    Top =1200
    Right =17190
    Bottom =11070
    DatasheetGridlinesColor =15132391
    OrderBy ="[frmTeams].[Teamnr]"
    RecSrcDt = Begin
        0x33fa47739292e540
    End
    RecordSource ="SELECT tblTeams.* FROM tblTeams WHERE (((tblTeams.ToernooiID)=lngToernooiID()));"
        " "
    Caption ="Teams"
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
                    Width =1386
                    Height =969
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Bijschrift24"
                    Caption ="Teams"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =1443
                    LayoutCachedHeight =1026
                End
            End
        End
        Begin Section
            Height =7740
            Name ="Details"
            AutoHeight =1
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
                    ColumnWidth =870
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
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =741
                    Width =1050
                    Height =330
                    ColumnWidth =3060
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Teamnr"
                    ControlSource ="Teamnr"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =741
                    LayoutCachedWidth =3942
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
                            Name ="Teamnr_Bijschrift"
                            Caption ="Teamnr"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =741
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =1071
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =1140
                    Width =7260
                    Height =600
                    ColumnWidth =1950
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TeamNaam"
                    ControlSource ="TeamNaam"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =1140
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =1740
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =1140
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="TeamNaam_Bijschrift"
                            Caption ="TeamNaam"
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
                    Top =1824
                    Width =7260
                    Height =600
                    ColumnWidth =1200
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Speler1"
                    ControlSource ="Speler1"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =1824
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =2424
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =1824
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Speler1_Bijschrift"
                            Caption ="Speler1"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =1824
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =2154
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =2508
                    Width =7260
                    Height =600
                    ColumnWidth =1425
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Speler2"
                    ControlSource ="Speler2"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =2508
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =3108
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =2508
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Speler2_Bijschrift"
                            Caption ="Speler2"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =2508
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =2838
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =3192
                    Width =7260
                    Height =600
                    ColumnWidth =1170
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Speler3"
                    ControlSource ="Speler3"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =3192
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =3792
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =3192
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Speler3_Bijschrift"
                            Caption ="Speler3"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =3192
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =3522
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =3876
                    Width =7260
                    Height =600
                    ColumnWidth =1410
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Speler4"
                    ControlSource ="Speler4"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =3876
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =4476
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =3876
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Speler4_Bijschrift"
                            Caption ="Speler4"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =3876
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =4206
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =4560
                    Width =7260
                    Height =600
                    ColumnWidth =990
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Speler5"
                    ControlSource ="Speler5"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =4560
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =5160
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =4560
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Speler5_Bijschrift"
                            Caption ="Speler5"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =4560
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =4890
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =5244
                    Width =7260
                    Height =600
                    ColumnWidth =1065
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Speler6"
                    ControlSource ="Speler6"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =5244
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =5844
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =5244
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Speler6_Bijschrift"
                            Caption ="Speler6"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =5244
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =5574
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =5928
                    Width =7260
                    Height =600
                    ColumnWidth =900
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Speler7"
                    ControlSource ="Speler7"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =5928
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =6528
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =5928
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Speler7_Bijschrift"
                            Caption ="Speler7"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =5928
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =6258
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =6612
                    Width =7260
                    Height =600
                    ColumnWidth =870
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Speler8"
                    ControlSource ="Speler8"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =6612
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =7212
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =6612
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Speler8_Bijschrift"
                            Caption ="Speler8"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =6612
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =6942
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =7296
                    Width =1530
                    Height =330
                    ColumnWidth =1530
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ToernooiID"
                    ControlSource ="ToernooiID"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =7296
                    LayoutCachedWidth =4422
                    LayoutCachedHeight =7626
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =342
                            Top =7296
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ToernooiID_Bijschrift"
                            Caption ="ToernooiID"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =7296
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =7626
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
