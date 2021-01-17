Version =20
VersionRequired =20
Begin Form
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8163
    DatasheetFontHeight =11
    ItemSuffix =47
    Right =19140
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
        Begin FormHeader
            Height =969
            BackColor =15064278
            Name ="Formulierkoptekst"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Width =3120
                    Height =969
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Bijschrift32"
                    Caption ="Process Viertallen"
                    GridlineColor =10921638
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =969
                End
            End
        End
        Begin Section
            Height =5116
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
                    Left =6066
                    Width =336
                    Height =315
                    ColumnOrder =0
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
                            OverlapFlags =85
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
                    OverlapFlags =93
                    Top =2267
                    Width =2091
                    Height =568
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnOphalenScores"
                    Caption ="--> Scorestaten"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedTop =2267
                    LayoutCachedWidth =2091
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
                    OverlapFlags =93
                    Top =3685
                    Width =2091
                    Height =568
                    TabIndex =2
                    ForeColor =4210752
                    Name ="btnHTMLUItslagen"
                    Caption ="--> HTML Uitslagen"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedTop =3685
                    LayoutCachedWidth =2091
                    LayoutCachedHeight =4253
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
                    OverlapFlags =215
                    Top =4252
                    Width =2076
                    Height =568
                    TabIndex =3
                    ForeColor =4210752
                    Name ="btnKruisTabel"
                    Caption ="--> HTML Kruistabel"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedTop =4252
                    LayoutCachedWidth =2076
                    LayoutCachedHeight =4820
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
                    OverlapFlags =93
                    Left =3005
                    Top =1756
                    ColumnOrder =1
                    TabIndex =4
                    BorderColor =10921638
                    Name ="optAlle"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3005
                    LayoutCachedTop =1756
                    LayoutCachedWidth =3265
                    LayoutCachedHeight =1996
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =5
                            Top =1700
                            Width =2660
                            Height =284
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblAlle"
                            Caption ="process alle scorekaarten"
                            GridlineColor =10921638
                            LayoutCachedLeft =5
                            LayoutCachedTop =1700
                            LayoutCachedWidth =2665
                            LayoutCachedHeight =1984
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =119
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3004
                    Top =1983
                    Width =3981
                    Height =315
                    ColumnOrder =2
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"3\";\"2\""
                    Name ="cboKiesTeam"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryTeams_per_Toernooi.Teamnr, qryTeams_per_Toernooi.TeamNaam FROM qryTeam"
                        "s_per_Toernooi WHERE (((qryTeams_per_Toernooi.ID)=lngToernooiID())) ORDER BY qry"
                        "Teams_per_Toernooi.Teamnr; "
                    ColumnWidths ="454;2835"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3004
                    LayoutCachedTop =1983
                    LayoutCachedWidth =6985
                    LayoutCachedHeight =2298
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =1
                            Left =5
                            Top =1983
                            Width =2660
                            Height =284
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblKiesTeam"
                            Caption ="process alleen team"
                            GridlineColor =10921638
                            LayoutCachedLeft =5
                            LayoutCachedTop =1983
                            LayoutCachedWidth =2665
                            LayoutCachedHeight =2267
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =177
                    Left =6803
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

                    LayoutCachedLeft =6803
                    LayoutCachedTop =3968
                    LayoutCachedWidth =7379
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
                    Top =3175
                    ColumnOrder =3
                    TabIndex =7
                    BorderColor =10921638
                    Name ="optHTML"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6803
                    LayoutCachedTop =3175
                    LayoutCachedWidth =7063
                    LayoutCachedHeight =3415
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =3968
                            Top =3118
                            Width =2495
                            Height =284
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblHTML"
                            Caption ="Uitvoer naar HTML"
                            GridlineColor =10921638
                            LayoutCachedLeft =3968
                            LayoutCachedTop =3118
                            LayoutCachedWidth =6463
                            LayoutCachedHeight =3402
                        End
                    End
                End
                Begin OptionButton
                    OverlapFlags =85
                    Left =6803
                    Top =623
                    ColumnOrder =4
                    TabIndex =8
                    BorderColor =10921638
                    Name ="optExcelZichtbaar"
                    OnClick ="[Event Procedure]"
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
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblExcelZichtbaar"
                            Caption ="Excelblad zichtbaar "
                            GridlineColor =10921638
                            LayoutCachedLeft =3968
                            LayoutCachedTop =566
                            LayoutCachedWidth =6463
                            LayoutCachedHeight =850
                        End
                    End
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OverlapFlags =215
                    Top =453
                    Width =3236
                    Height =1273
                    TabIndex =9
                    BorderColor =10921638
                    Name ="grpUitvoernaar"
                    DefaultValue ="1"
                    GridlineColor =10921638

                    LayoutCachedTop =453
                    LayoutCachedWidth =3236
                    LayoutCachedHeight =1726
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =11
                            Top =453
                            Width =1245
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblUitvoerNaar"
                            Caption ="Uitvoer Naar"
                            GridlineColor =10921638
                            LayoutCachedLeft =11
                            LayoutCachedTop =453
                            LayoutCachedWidth =1256
                            LayoutCachedHeight =768
                            BackThemeColorIndex =-1
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =2976
                            Top =691
                            OptionValue =1
                            BorderColor =10921638
                            Name ="optExcel"
                            GridlineColor =10921638

                            LayoutCachedLeft =2976
                            LayoutCachedTop =691
                            LayoutCachedWidth =3236
                            LayoutCachedHeight =931
                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =1361
                                    Top =663
                                    Width =1134
                                    Height =315
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="Bijschrift39"
                                    Caption ="Excel"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =1361
                                    LayoutCachedTop =663
                                    LayoutCachedWidth =2495
                                    LayoutCachedHeight =978
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =2976
                            Top =1021
                            TabIndex =1
                            OptionValue =2
                            BorderColor =10921638
                            Name ="optAccess"
                            GridlineColor =10921638

                            LayoutCachedLeft =2976
                            LayoutCachedTop =1021
                            LayoutCachedWidth =3236
                            LayoutCachedHeight =1261
                            Begin
                                Begin Label
                                    OverlapFlags =95
                                    Left =1361
                                    Top =993
                                    Width =1134
                                    Height =315
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="Bijschrift41"
                                    Caption ="Access"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =1361
                                    LayoutCachedTop =993
                                    LayoutCachedWidth =2495
                                    LayoutCachedHeight =1308
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =2976
                            Top =1351
                            TabIndex =2
                            OptionValue =3
                            BorderColor =10921638
                            Name ="OptBeiden"
                            GridlineColor =10921638

                            LayoutCachedLeft =2976
                            LayoutCachedTop =1351
                            LayoutCachedWidth =3236
                            LayoutCachedHeight =1591
                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =1361
                                    Top =1304
                                    Width =1134
                                    Height =315
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="Bijschrift43"
                                    Caption ="Beiden"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =1361
                                    LayoutCachedTop =1304
                                    LayoutCachedWidth =2495
                                    LayoutCachedHeight =1619
                                End
                            End
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Width =2400
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblBerekenScorekaarten"
                    Caption ="Berekening Scorekaarten"
                    GridlineColor =10921638
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =315
                End
                Begin Label
                    OverlapFlags =85
                    Top =3118
                    Width =1740
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Bijschrift45"
                    Caption ="Extra html uitvoer"
                    GridlineColor =10921638
                    LayoutCachedTop =3118
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =3433
                End
                Begin Label
                    OverlapFlags =85
                    Left =2993
                    Top =2381
                    Width =4995
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lbGekozenTeam"
                    GridlineColor =10921638
                    LayoutCachedLeft =2993
                    LayoutCachedTop =2381
                    LayoutCachedWidth =7988
                    LayoutCachedHeight =2696
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
        x = VulScoreKaartInSheet(CInt(Me.cboKiesTeam), Sessienr, 2, lngToernooi, True, False)
        DoCmd.OpenForm "frmScorestaat", acNormal, , "[ToernooiID] = " & lngToernooi & " and [id] = " & WijID
    Else
        If Me.optAlle = True Then
            Call AlleScoreStaten_RESULTS(Sessienr, lngToernooi, lngSessie)
        End If
    End If


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
