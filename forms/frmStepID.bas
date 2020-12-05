Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =13140
    DatasheetFontHeight =11
    ItemSuffix =11
    Right =13995
    Bottom =10470
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
            Height =8503
            Name ="Details"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin ComboBox
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
End Sub
