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
    Width =10204
    DatasheetFontHeight =11
    ItemSuffix =30
    Left =570
    Top =1290
    Right =16995
    Bottom =9315
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0xe10fa2fe7892e540
    End
    RecordSource ="SELECT tblScorestaat.* FROM tblScorestaat WHERE (((tblScorestaat.SessieID)=lngSe"
        "ssieID())); "
    Caption ="tblScorestaat"
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
                    Width =2664
                    Height =969
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Bijschrift26"
                    Caption ="tblScorestaat"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =2721
                    LayoutCachedHeight =1026
                End
            End
        End
        Begin Section
            Height =6999
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
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Spelnr"
                    ControlSource ="Spelnr"
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
                            Name ="Spelnr_Bijschrift"
                            Caption ="Spelnr"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =342
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =672
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =741
                    Width =7260
                    Height =600
                    ColumnWidth =2460
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Contract1"
                    ControlSource ="Contract1"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =741
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =1341
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =741
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Contract1_Bijschrift"
                            Caption ="Contract1"
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
                    Top =1425
                    Width =7260
                    Height =600
                    ColumnWidth =630
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Resultaat1"
                    ControlSource ="Resultaat1"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =1425
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =2025
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =1425
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Resultaat1_Bijschrift"
                            Caption ="Resultaat1"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =1425
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =1755
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =2109
                    Width =840
                    Height =330
                    ColumnWidth =645
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Door1"
                    ControlSource ="Door1"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =2109
                    LayoutCachedWidth =3732
                    LayoutCachedHeight =2439
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =2109
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Door1_Bijschrift"
                            Caption ="Door1"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =2109
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =2439
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =2508
                    Width =1050
                    Height =330
                    ColumnWidth =1050
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Score1"
                    ControlSource ="Score1"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =2508
                    LayoutCachedWidth =3942
                    LayoutCachedHeight =2838
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =2508
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Score1_Bijschrift"
                            Caption ="Score1"
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
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2892
                    Top =2907
                    Width =7260
                    Height =600
                    ColumnWidth =2220
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Contract2"
                    ControlSource ="Contract2"
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
                            Name ="Contract2_Bijschrift"
                            Caption ="Contract2"
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
                    ColumnWidth =855
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Resultaat2"
                    ControlSource ="Resultaat2"
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
                            Name ="Resultaat2_Bijschrift"
                            Caption ="Resultaat2"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =3591
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =3921
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =4275
                    Width =7260
                    Height =600
                    ColumnWidth =735
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Door2"
                    ControlSource ="Door2"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =4275
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =4875
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =4275
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Door2_Bijschrift"
                            Caption ="Door2"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =4275
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =4605
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =4959
                    Width =1050
                    Height =330
                    ColumnWidth =0
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Score2"
                    ControlSource ="Score2"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =4959
                    LayoutCachedWidth =3942
                    LayoutCachedHeight =5289
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =4959
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Score2_Bijschrift"
                            Caption ="Score2"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =4959
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =5289
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =5358
                    Width =1050
                    Height =330
                    ColumnWidth =1005
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Saldo"
                    ControlSource ="Saldo"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =5358
                    LayoutCachedWidth =3942
                    LayoutCachedHeight =5688
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =5358
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Saldo_Bijschrift"
                            Caption ="Saldo"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =5358
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =5688
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =5757
                    Width =1050
                    Height =330
                    ColumnWidth =1050
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Imps"
                    ControlSource ="Imps"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =5757
                    LayoutCachedWidth =3942
                    LayoutCachedHeight =6087
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =5757
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Imps_Bijschrift"
                            Caption ="Imps"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =5757
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =6087
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =6156
                    Width =1050
                    Height =330
                    ColumnWidth =1050
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="WijImps"
                    ControlSource ="WijImps"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =6156
                    LayoutCachedWidth =3942
                    LayoutCachedHeight =6486
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =6156
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="WijImps_Bijschrift"
                            Caption ="WijImps"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =6156
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =6486
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =6555
                    Width =1050
                    Height =330
                    ColumnWidth =1050
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ZijImps"
                    ControlSource ="ZijImps"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =6555
                    LayoutCachedWidth =3942
                    LayoutCachedHeight =6885
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =6555
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ZijImps_Bijschrift"
                            Caption ="ZijImps"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =6555
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =6885
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5725
                    Top =113
                    Height =315
                    TabIndex =13
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ToernooiID"
                    ControlSource ="ToernooiID"
                    GridlineColor =10921638

                    LayoutCachedLeft =5725
                    LayoutCachedTop =113
                    LayoutCachedWidth =7426
                    LayoutCachedHeight =428
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =4024
                            Top =113
                            Width =1110
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift27"
                            Caption ="ToernooiID"
                            GridlineColor =10921638
                            LayoutCachedLeft =4024
                            LayoutCachedTop =113
                            LayoutCachedWidth =5134
                            LayoutCachedHeight =428
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7086
                    Top =2154
                    Height =315
                    TabIndex =14
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TeamID"
                    ControlSource ="TeamID"
                    GridlineColor =10921638

                    LayoutCachedLeft =7086
                    LayoutCachedTop =2154
                    LayoutCachedWidth =8787
                    LayoutCachedHeight =2469
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =5385
                            Top =2154
                            Width =780
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift28"
                            Caption ="TeamID"
                            GridlineColor =10921638
                            LayoutCachedLeft =5385
                            LayoutCachedTop =2154
                            LayoutCachedWidth =6165
                            LayoutCachedHeight =2469
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =6179
                    Top =2607
                    Height =315
                    TabIndex =15
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="SessieID"
                    ControlSource ="SessieID"
                    GridlineColor =10921638

                    LayoutCachedLeft =6179
                    LayoutCachedTop =2607
                    LayoutCachedWidth =7880
                    LayoutCachedHeight =2922
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =247
                            Left =4478
                            Top =2607
                            Width =855
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift29"
                            Caption ="SessieID"
                            GridlineColor =10921638
                            LayoutCachedLeft =4478
                            LayoutCachedTop =2607
                            LayoutCachedWidth =5333
                            LayoutCachedHeight =2922
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
