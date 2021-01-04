Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10209
    DatasheetFontHeight =11
    ItemSuffix =19
    Left =2805
    Top =2340
    Right =14385
    Bottom =7515
    DatasheetGridlinesColor =15132391
    Filter ="[WedstrijdRonde] = 7"
    RecSrcDt = Begin
        0x6c8ed59ad094e540
    End
    RecordSource ="SELECT tblSchema.*, tblTeams.TeamNaam AS TeamnaamThuis, tblTeams_1.TeamNaam AS T"
        "eamNaamUit FROM (tblSchema INNER JOIN tblTeams ON (tblSchema.ToernooiID = tblTea"
        "ms.ToernooiID) AND (tblSchema.TeamThuis = tblTeams.Teamnr)) INNER JOIN tblTeams "
        "AS tblTeams_1 ON (tblSchema.ToernooiID = tblTeams_1.ToernooiID) AND (tblSchema.T"
        "eamUit = tblTeams_1.Teamnr) WHERE (((tblSchema.ToernooiID)=lngToernooiID())) ORD"
        "ER BY tblSchema.ToernooiID, tblSchema.Wedstrijdronde, tblSchema.Paring; "
    Caption ="frmSchema"
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
                    Width =2322
                    Height =969
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Bijschrift16"
                    Caption ="frmSchema"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =2379
                    LayoutCachedHeight =1026
                End
            End
        End
        Begin Section
            Height =4149
            Name ="Details"
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =342
                    Width =1050
                    Height =330
                    ColumnWidth =1050
                    ColumnOrder =0
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Wedstrijdronde"
                    ControlSource ="Wedstrijdronde"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =342
                    LayoutCachedWidth =3942
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
                            Name ="Wedstrijdronde_Bijschrift"
                            Caption ="Wedstrijdronde"
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
                    ColumnWidth =1050
                    ColumnOrder =1
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Paring"
                    ControlSource ="Paring"
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
                            Name ="Paring_Bijschrift"
                            Caption ="Paring"
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
                    ColumnOrder =2
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TeamThuis"
                    ControlSource ="TeamThuis"
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
                            Name ="TeamThuis_Bijschrift"
                            Caption ="TeamThuis"
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
                    ColumnWidth =1050
                    ColumnOrder =4
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TeamUit"
                    ControlSource ="TeamUit"
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
                            Name ="TeamUit_Bijschrift"
                            Caption ="TeamUit"
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
                    ColumnWidth =1095
                    ColumnOrder =6
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Tafel1"
                    ControlSource ="Tafel1"
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
                            Name ="Tafel1_Bijschrift"
                            Caption ="Tafel1"
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
                    ColumnWidth =1530
                    ColumnOrder =7
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Tafel2"
                    ControlSource ="Tafel2"
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
                            Name ="Tafel2_Bijschrift"
                            Caption ="Tafel2"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =2622
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =2952
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =3306
                    Width =1530
                    Height =330
                    ColumnWidth =1530
                    ColumnOrder =8
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ToernooiID"
                    ControlSource ="ToernooiID"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =3306
                    LayoutCachedWidth =4422
                    LayoutCachedHeight =3636
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =342
                            Top =3306
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ToernooiID_Bijschrift"
                            Caption ="ToernooiID"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =3306
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =3636
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =3705
                    Height =315
                    ColumnWidth =1701
                    ColumnOrder =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Id"
                    ControlSource ="Id"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =3705
                    LayoutCachedWidth =4593
                    LayoutCachedHeight =4020
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =342
                            Top =3705
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Id_Bijschrift"
                            Caption ="Id"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =3705
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =4035
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7596
                    Top =963
                    Height =315
                    ColumnWidth =2685
                    ColumnOrder =3
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TeamnaamThuis"
                    ControlSource ="TeamnaamThuis"
                    GridlineColor =10921638

                    LayoutCachedLeft =7596
                    LayoutCachedTop =963
                    LayoutCachedWidth =9297
                    LayoutCachedHeight =1278
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5895
                            Top =963
                            Width =1590
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift17"
                            Caption ="TeamnaamThuis"
                            GridlineColor =10921638
                            LayoutCachedLeft =5895
                            LayoutCachedTop =963
                            LayoutCachedWidth =7485
                            LayoutCachedHeight =1278
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7256
                    Top =1530
                    Height =315
                    ColumnWidth =1710
                    ColumnOrder =5
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TeamNaamUit"
                    ControlSource ="TeamNaamUit"
                    GridlineColor =10921638

                    LayoutCachedLeft =7256
                    LayoutCachedTop =1530
                    LayoutCachedWidth =8957
                    LayoutCachedHeight =1845
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5555
                            Top =1530
                            Width =1395
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift18"
                            Caption ="TeamNaamUit"
                            GridlineColor =10921638
                            LayoutCachedLeft =5555
                            LayoutCachedTop =1530
                            LayoutCachedWidth =6950
                            LayoutCachedHeight =1845
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
