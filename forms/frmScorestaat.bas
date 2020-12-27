Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =17291
    DatasheetFontHeight =11
    ItemSuffix =12
    Right =18090
    Bottom =10245
    DatasheetGridlinesColor =15132391
    Filter ="[ToernooiID] = 1 and [id] = 13"
    RecSrcDt = Begin
        0x9436a4437a92e540
    End
    RecordSource ="SELECT tblTeams.* FROM tblTeams WHERE (((tblTeams.ToernooiID)=lngToernooiID()));"
        " "
    OnOpen ="[Event Procedure]"
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
            CanGrow = NotDefault
            Height =1052
            BackColor =15064278
            Name ="Formulierkoptekst"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin ComboBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =2834
                    Width =3981
                    Height =315
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"3\";\"2\""
                    Name ="cboKiesTeam"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblTeams.ID, tblTeams.Teamnr, tblTeams.TeamNaam FROM tblTeams WHERE (((tb"
                        "lTeams.ToernooiID)=lngToernooiID())) ORDER BY tblTeams.Teamnr; "
                    ColumnWidths ="0;456;2835"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =2834
                    LayoutCachedWidth =6815
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Width =2495
                            Height =284
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblKiesTeam"
                            Caption ="Team"
                            GridlineColor =10921638
                            LayoutCachedWidth =2495
                            LayoutCachedHeight =284
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14286
                    Top =113
                    Width =246
                    Height =315
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="id"
                    ControlSource ="id"
                    GridlineColor =10921638

                    LayoutCachedLeft =14286
                    LayoutCachedTop =113
                    LayoutCachedWidth =14532
                    LayoutCachedHeight =428
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =13719
                            Top =170
                            Width =255
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift1"
                            Caption ="id"
                            GridlineColor =10921638
                            LayoutCachedLeft =13719
                            LayoutCachedTop =170
                            LayoutCachedWidth =13974
                            LayoutCachedHeight =485
                        End
                    End
                End
                Begin ComboBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =2834
                    Top =510
                    Width =3981
                    Height =315
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"3\";\"2\""
                    Name ="cboKiesSessie"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblSessie.id, tblSessie.Sessienr FROM tblSessie WHERE (((tblSessie.Toerno"
                        "oID)=lngToernooiID())) ORDER BY tblSessie.[Sessienr]; "
                    ColumnWidths ="0;1440"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =2834
                    LayoutCachedTop =510
                    LayoutCachedWidth =6815
                    LayoutCachedHeight =825
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =53
                            Top =510
                            Width =2430
                            Height =320
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Kies Sessie_Etiket"
                            Caption ="Kies Sessie"
                            EventProcPrefix ="Kies_Sessie_Etiket"
                            GridlineColor =10921638
                            LayoutCachedLeft =53
                            LayoutCachedTop =510
                            LayoutCachedWidth =2483
                            LayoutCachedHeight =830
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =8560
            Name ="Details"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Subform
                    OverlapFlags =215
                    Left =283
                    Top =240
                    Width =16440
                    Height =8295
                    BorderColor =10921638
                    Name ="tblScorestaat"
                    SourceObject ="Form.tblScorestaat"
                    LinkChildFields ="TeamID"
                    LinkMasterFields ="id"
                    GridlineColor =10921638

                    LayoutCachedLeft =283
                    LayoutCachedTop =240
                    LayoutCachedWidth =16723
                    LayoutCachedHeight =8535
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =283
                            Width =1290
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift0"
                            Caption ="tblScorestaat"
                            GridlineColor =10921638
                            LayoutCachedLeft =283
                            LayoutCachedWidth =1573
                            LayoutCachedHeight =315
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
Dim MyKey As String
Dim MyKeyIs As String

Private Sub cboKiesSessie_AfterUpdate()
  lngSessie = Me.cboKiesSessie
  Call InitAll(lngToernooi, lngSessie)
  Me.tblScorestaat.Requery
 
End Sub

Private Sub cboKiesTeam_AfterUpdate()
Dim rs As Recordset
    Dim x
   Dim Criterium As String
  ' If Me.Dirty Then
   '   x = fnSaveRecords
   ' End If
    Set rs = Me.Recordset
   
    Criterium = MyKeyIs & cboKiesTeam
    rs.FindFirst Criterium
    
    If rs.NoMatch Then
    MsgBox "Geen Team bekend in de database"
        cboKiesTeam = ""
    Else
    Me.Bookmark = rs.Bookmark
        cboKiesTeam = ""
    End If
    
End Sub

Private Sub Form_Open(Cancel As Integer)
MyKey = "ID"
MyKeyIs = MyKey & " = "
Me.cboKiesTeam = ""
If lngToernooi = 0 Then
    lngToernooi = 1
    lngSessie = 1
    Call InitAll(lngToernooi, lngSessie)
End If
If CurrentProject.AllForms("Start_VT").IsLoaded = True Then
    Me.cboKiesTeam.Visible = False
    Me.cboKiesTeam.Enabled = False
    Me.cboKiesSessie.Visible = False
    Me.cboKiesSessie.Enabled = False
    
    
    
Else
    Me.cboKiesTeam.Visible = True
    Me.cboKiesTeam.Enabled = True
    Me.cboKiesSessie.Visible = True
    Me.cboKiesSessie.Enabled = True

End If

End Sub
