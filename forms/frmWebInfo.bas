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
    Width =4650
    DatasheetFontHeight =11
    ItemSuffix =11
    Right =18990
    Bottom =12240
    DatasheetGridlinesColor =15132391
    OrderBy ="[tblWebInfo].[ToernooiID], [tblWebInfo].[Sessie]"
    RecSrcDt = Begin
        0x934ec1183190e540
    End
    RecordSource ="tblWebInfo"
    Caption ="frmWebInfo"
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
            Height =1077
            BackColor =15064278
            Name ="Formulierkoptekst"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Width =2484
                    Height =969
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Bijschrift8"
                    Caption ="frmWebInfo"
                    GridlineColor =10921638
                    LayoutCachedWidth =2484
                    LayoutCachedHeight =969
                End
            End
        End
        Begin Section
            Height =3798
            Name ="Details"
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
                    ColumnWidth =1701
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
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =741
                    Width =1530
                    Height =330
                    ColumnWidth =1530
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
                    Width =1050
                    Height =330
                    ColumnWidth =1050
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Sessie"
                    ControlSource ="Sessie"
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
                    Width =1530
                    Height =330
                    ColumnWidth =1530
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ActivityID"
                    ControlSource ="ActivityID"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =1539
                    LayoutCachedWidth =4422
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
                            Name ="ActivityID_Bijschrift"
                            Caption ="ActivityID"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =1539
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =1869
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2211
                    Top =2211
                    Height =315
                    ColumnWidth =4320
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="Keuzelijst9"
                    ControlSource ="ToernooiID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblToernooi.ID, tblToernooi.ToernooiNaam FROM tblToernooi ORDER BY tblToe"
                        "rnooi.ToernooiNaam; "
                    ColumnWidths ="0;2835"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =2211
                    LayoutCachedTop =2211
                    LayoutCachedWidth =3912
                    LayoutCachedHeight =2526
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =510
                            Top =2211
                            Width =1110
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift10"
                            Caption ="ToernooiID"
                            GridlineColor =10921638
                            LayoutCachedLeft =510
                            LayoutCachedTop =2211
                            LayoutCachedWidth =1620
                            LayoutCachedHeight =2526
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

Private Sub ActivityID_AfterUpdate()
'fris sessie tabel op


End Sub
