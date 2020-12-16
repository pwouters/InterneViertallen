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
    Width =10780
    DatasheetFontHeight =11
    ItemSuffix =33
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
            Height =1134
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
                    Width =3120
                    Height =969
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Bijschrift32"
                    Caption ="Kies Toernooi"
                    GridlineColor =10921638
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =969
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =5387
                    Top =170
                    Width =5376
                    Height =315
                    ColumnOrder =0
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="cboKiesToernooi"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [tblToernooi].[ID], [tblToernooi].[ToernooiNaam] FROM tblToernooi ORDER B"
                        "Y [ToernooiNaam]; "
                    ColumnWidths ="0;1440"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =5387
                    LayoutCachedTop =170
                    LayoutCachedWidth =10763
                    LayoutCachedHeight =485
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3401
                            Top =170
                            Width =1680
                            Height =320
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Kies Toernooi_Etiket"
                            Caption ="Kies Toernooi"
                            EventProcPrefix ="Kies_Toernooi_Etiket"
                            GridlineColor =10921638
                            LayoutCachedLeft =3401
                            LayoutCachedTop =170
                            LayoutCachedWidth =5081
                            LayoutCachedHeight =490
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2760
                    Left =5387
                    Top =566
                    Width =5376
                    Height =315
                    ColumnOrder =1
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"3\";\"2\""
                    Name ="cboKiesSessie"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tblSessie.Sessienr, tblSessie.Sessienaam FROM tblSessie WHERE (((tblSessi"
                        "e.ToernooID)=lngToernooiID())) ORDER BY tblSessie.Sessienr, tblSessie.Sessienaam"
                        "; "
                    ColumnWidths ="567;2835"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    GridlineColor =10921638

                    LayoutCachedLeft =5387
                    LayoutCachedTop =566
                    LayoutCachedWidth =10763
                    LayoutCachedHeight =881
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3401
                            Top =566
                            Width =1695
                            Height =320
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Sessienr_Etiket"
                            Caption ="Sessienr"
                            GridlineColor =10921638
                            LayoutCachedLeft =3401
                            LayoutCachedTop =566
                            LayoutCachedWidth =5096
                            LayoutCachedHeight =886
                        End
                    End
                End
            End
        End
        Begin Section
            Height =2324
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
                    Left =1814
                    Height =315
                    ColumnOrder =0
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ID"
                    ControlSource ="ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =1814
                    LayoutCachedWidth =3515
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =1360
                            Width =270
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift4"
                            Caption ="ID"
                            GridlineColor =10921638
                            LayoutCachedLeft =1360
                            LayoutCachedWidth =1630
                            LayoutCachedHeight =315
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =177
                    Left =3968
                    Top =1133
                    Width =576
                    Height =576
                    TabIndex =1
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

                    LayoutCachedLeft =3968
                    LayoutCachedTop =1133
                    LayoutCachedWidth =4544
                    LayoutCachedHeight =1709
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
                Begin CommandButton
                    OverlapFlags =93
                    Left =566
                    Top =566
                    Width =2106
                    Height =568
                    TabIndex =2
                    ForeColor =4210752
                    Name ="btnNewToernooi"
                    Caption ="Nieuw Toernooi"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =566
                    LayoutCachedWidth =2672
                    LayoutCachedHeight =1134
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
                    OverlapFlags =87
                    Left =566
                    Top =1133
                    Width =2106
                    Height =568
                    TabIndex =3
                    ForeColor =4210752
                    Name ="btnNieuweSessie"
                    Caption ="Nieuwe Sessie"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =1133
                    LayoutCachedWidth =2672
                    LayoutCachedHeight =1701
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

Private Sub btnHTMLUItslagen_Click()
Call HTMLViertalUitslagenIn(CInt(Me.cboKiesSessie), CInt(Me.cboKiesToernooi), lngSessie)
End Sub



Private Sub btnKruisTabel_Click()
Call HTMLViertalKruistabel(CInt(Me.cboKiesToernooi))
End Sub

Private Sub btnNewToernooi_Click()
' Kies worktemplate
If Not KiesModelExcelBestandEnCopieer Then
    MsgBox ("Process nieuw toernooi is afgebroken")
    Exit Sub
End If
Dim db As Database
Dim rs As Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblToernooi")
    rs.AddNew
    rs!ToernooiNaam = ToernooiNaam
    rs!WORKFOLDER = WORKFOLDER
    rs!WORKFILE = WORKFILE
    rs!STEPDATA = STEP_DATA
    rs!STEPRESULTS = STEP_RESULTS
    rs!AANTALSESSIES = 1
    rs!WEDSTRIJDENPERSESSIE = 1
    rs!WORKTEMPLATE = WORKFOLDER
    rs!LOCALHTML = LOCALHTML
    rs!LOCALHTML = LOCALSITE
    rs!PREFIX = "Wedstrijd_"
    rs.Update
    rs.Bookmark = rs.LastModified
    lngToernooi = rs!Id
    rs.Close
    
    'nu add sessie aan
    Set rs = db.OpenRecordset("tblSessie")
    
    rs.AddNew
    rs!ToernooID = lngToernooi
    rs!Sessienaam = "Eerste Sessie"
    rs!Sessienr = 1
    rs!Aantalspellen = 8
    rs!Competitie = 1  'halve
    rs!Prefixkopjesscorestaat = "Scorekaart van sessie 1 "
    rs!PrefixKopjeuitslagen = "Uitslag van sessie 1"
    rs!Suffixkopjesscorestaat = ""
    rs!SuffixKopjeuitslagen = ""
    rs!Voettekst = "@Bridge"
    rs!Voetlink = "#"
    rs!wedstrijdvormID = 0
    rs!ActivityID = 0
    rs!AANTALTEAMS = 16
    rs!ByeTeam = False
    rs!AantalWedstrijdenPerSessie = 1
    rs.Update
    rs.Bookmark = rs.LastModified
    lngSessie = rs!Id
    
rs.Close
db.Close

    Call InitAll(lngToernooi, lngSessie)
     
     Me.cboKiesToernooi.Requery
     Me.cboKiesToernooi = lngToernooi
     Me.cboKiesSessie.Requery
     Me.cboKiesSessie = Sessienr
     

End Sub

Private Sub btnNieuweSessie_Click()
Dim intSessie As Integer
Dim x
    'hoog sessie nr op
    'dupliceer gegevens van de vorige sessie
    intSessie = DMax("Sessienr", "tblSessie", "[ToernooiD]= " & lngToernooi)
    lngSessie = DLookup("id", "tblSessie", "[ToernooiD]= " & lngToernooi & " and [Sessienr] = " & intSessie)
    Call InitAll(lngToernooi, lngSessie)
    
   If intSessie = AANTALSESSIES Then
   
     x = MsgBox("(" & ToernooiNaam & ")" & " Je hebt al het laatste sessie nr bereikt, je kunt dit alleen ophogen bij toernooigegevens")
     Exit Sub
   End If
   
   intSessie = intSessie + 1
    x = MsgBox("(" & ToernooiNaam & ")" & " nieuwe sessie nr " & intSessie & " toevoegen? J/N", vbYesNo)
   If x <> vbYes Then
        Exit Sub
   End If
   
Dim db As Database
Dim rs As Recordset
Set db = CurrentDb
Set rs = db.OpenRecordset("tblSessie")
  rs.AddNew
    rs!ToernooID = lngToernooi
    rs!Sessienaam = Sessienaam
    
    rs!Sessienr = intSessie
    rs!Aantalspellen = AANTALSPELLENPERWEDSTRIJD
    rs!Competitie = WEDSTRIJD
    rs!Prefixkopjesscorestaat = Prefixkopjesscorestaat
    rs!PrefixKopjeuitslagen = PrefixKopjeuitslagen
    rs!Suffixkopjesscorestaat = Suffixkopjesscorestaat
    rs!SuffixKopjeuitslagen = SuffixKopjeuitslagen
    rs!Voettekst = Voettekst
    rs!Voetlink = Voetlink
    rs!wedstrijdvormID = COMPETITIEVORM
    rs!ActivityID = 0
    rs!AANTALTEAMS = AANTALTEAMS
    If TEAMBYE = AANTALTEAMS Then
        rs!ByeTeam = True
    Else
        rs!ByeTeam = False
    End If
    rs!AantalWedstrijdenPerSessie = WEDSTRIJDENPERSESSIE
    rs.Update
    rs.Bookmark = rs.LastModified
    lngSessie = rs!Id
    
rs.Close
db.Close

    Call InitAll(lngToernooi, lngSessie)
    
     Me.cboKiesToernooi.Requery
     Me.cboKiesToernooi = lngToernooi
     Me.cboKiesSessie.Requery
     Me.cboKiesSessie = Sessienr
    
    
End Sub





Private Sub btnSluiten_Click()
    DoCmd.Close
End Sub

Private Sub cboKiesSessie_AfterUpdate()
     lngSessie = DLookup("id", "tblSessie", "Sessienr=" & Me.cboKiesSessie & " and ToernooiD = " & lngToernooi)
     Call InitAll(lngToernooi, lngSessie)
  
End Sub

Private Sub cboKiesToernooi_AfterUpdate()
     lngToernooi = Me.cboKiesToernooi
    'lngSessie = DLookup("id", "tblSessie", "Sessienr=" & Me.cboKiesSessie & " and ToernooiD = " & lngToernooi)
     lngSessie = DLookup("id", "tblSessie", "Sessienr=" & 1 & " and ToernooiD = " & lngToernooi)
    
     Call InitAll(lngToernooi, lngSessie)
    
     Me.cboKiesSessie.Requery
     Me.cboKiesSessie = Sessienr
   
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
   
   Me.cboKiesSessie.Requery
   Me.cboKiesSessie.Value = Sessienr
   Me.cboKiesToernooi.Value = lngToernooi



    

End Sub
