Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11508
    DatasheetFontHeight =11
    ItemSuffix =11
    Right =18735
    Bottom =12240
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0x8832de423290e540
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
        Begin EmptyCell
            Height =240
            GridlineColor =12632256
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            CanGrow = NotDefault
            Height =7483
            Name ="Details"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =2441
                    Top =566
                    Width =5106
                    Height =315
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="cboKeuzelijst"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [tblToernooi].[ID], [tblToernooi].[ToernooiNaam] FROM tblToernooi ORDER B"
                        "Y [ToernooiNaam]; "
                    ColumnWidths ="0;1440"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =2441
                    LayoutCachedTop =566
                    LayoutCachedWidth =7547
                    LayoutCachedHeight =881
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =170
                            Top =566
                            Width =1995
                            Height =320
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Kies Toernooi_Etiket"
                            Caption ="Kies Toernooi"
                            EventProcPrefix ="Kies_Toernooi_Etiket"
                            GridlineColor =10921638
                            LayoutCachedLeft =170
                            LayoutCachedTop =566
                            LayoutCachedWidth =2165
                            LayoutCachedHeight =886
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =215
                    Left =625
                    Top =2834
                    Width =10140
                    Height =3270
                    TabIndex =1
                    BorderColor =10921638
                    Name ="frmWebInfo"
                    SourceObject ="Form.frmWebInfo"
                    LinkChildFields ="ToernooiID"
                    LinkMasterFields ="ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =625
                    LayoutCachedTop =2834
                    LayoutCachedWidth =10765
                    LayoutCachedHeight =6104
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1870
                            Top =2594
                            Width =1215
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift9"
                            Caption ="frmWebInfo"
                            GridlineColor =10921638
                            LayoutCachedLeft =1870
                            LayoutCachedTop =2594
                            LayoutCachedWidth =3085
                            LayoutCachedHeight =2909
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =9929
                    Top =623
                    Width =501
                    Height =510
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ID"
                    ControlSource ="ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =9929
                    LayoutCachedTop =623
                    LayoutCachedWidth =10430
                    LayoutCachedHeight =1133
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =9589
                            Top =623
                            Width =375
                            Height =510
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift10"
                            Caption ="ID"
                            GridlineColor =10921638
                            LayoutCachedLeft =9589
                            LayoutCachedTop =623
                            LayoutCachedWidth =9964
                            LayoutCachedHeight =1133
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =177
                    Left =10771
                    Top =6803
                    Width =576
                    Height =576
                    TabIndex =3
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
                    LayoutCachedTop =6803
                    LayoutCachedWidth =11347
                    LayoutCachedHeight =7379
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
Public MyKey As String
Public MyKeyIs  As String



Private Sub btnSluiten_Click()
    If CurrentProject.AllForms("Start_VT").IsLoaded = False Then
        DoCmd.Close
    Else
        DoCmd.BrowseTo acBrowseToForm, "frmProcess"
   End If
End Sub

Private Sub cboKeuzelijst_AfterUpdate()
Dim rs As Recordset
    Dim X
   Dim criterium As String
   If Me.Dirty Then
      ' x = fnSaveRecords
    End If
    Set rs = Recordset
   
    criterium = MyKeyIs & cboKeuzelijst
    rs.FindFirst criterium
    
    If rs.NoMatch Then
    MsgBox "Geen lijst bekend in de database"
    cboKeuzelijst = ""
    Else
    Me.Bookmark = rs.Bookmark
    cboKeuzelijst = ""
    End If
End Sub



Private Sub Form_Open(Cancel As Integer)
MyKey = "ID"
MyKeyIs = MyKey & " = "
Me.cboKeuzelijst = ""
'indien geen navigatie dan zichtbaar anders verbergen en filter

If CurrentProject.AllForms("Start_VT").IsLoaded = False Then
        Me.cboKeuzelijst.Visible = True
        Me.cboKeuzelijst.Enabled = True
        Else
        Me.Filter = "[id]=" & lngToernooi
        Me.FilterOn = True
End If




End Sub
