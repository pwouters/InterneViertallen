Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11914
    DatasheetFontHeight =11
    ItemSuffix =27
    Right =18885
    Bottom =12240
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0xada97b1eb092e540
    End
    RecordSource ="SELECT tblToernooi.ID, tblToernooi.ToernooiNaam, tblToernooi.AANTALSESSIES FROM "
        "tblToernooi WHERE (((tblToernooi.ID)=lngToernooiID())); "
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
                    Left =566
                    Top =226
                    Width =3630
                    Height =570
                    FontSize =18
                    BorderColor =8355711
                    ForeColor =-2147483617
                    Name ="Bijschrift0"
                    Caption ="Schema verwerking"
                    GridlineColor =10921638
                    LayoutCachedLeft =566
                    LayoutCachedTop =226
                    LayoutCachedWidth =4196
                    LayoutCachedHeight =796
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =2880
                    Left =6336
                    Top =340
                    Width =3516
                    Height =315
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="cboKiesRonde"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT tblSchema.Wedstrijdronde FROM tblSchema WHERE (((tblSchema.Toern"
                        "ooiID)=lngToernooiID())) union ALL SELECT TOP 1 'Alle'  FROM tblSchema"
                    ColumnWidths ="1442"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6336
                    LayoutCachedTop =340
                    LayoutCachedWidth =9852
                    LayoutCachedHeight =655
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4365
                            Top =340
                            Width =1920
                            Height =320
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Kies Ronde_Etiket"
                            Caption ="Kies Ronde"
                            EventProcPrefix ="Kies_Ronde_Etiket"
                            GridlineColor =10921638
                            LayoutCachedLeft =4365
                            LayoutCachedTop =340
                            LayoutCachedWidth =6285
                            LayoutCachedHeight =660
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =7937
            Name ="Details"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    Left =5669
                    Top =6236
                    Width =2271
                    Height =568
                    ForeColor =4210752
                    Name ="btnSchemaNaarUitslag"
                    Caption ="Schema --> Uitslag"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =5669
                    LayoutCachedTop =6236
                    LayoutCachedWidth =7940
                    LayoutCachedHeight =6804
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
                Begin OptionGroup
                    OverlapFlags =85
                    Left =9127
                    Top =6242
                    Width =1247
                    Height =943
                    TabIndex =1
                    BorderColor =10921638
                    Name ="grpUitvoerNaar"
                    DefaultValue ="1"
                    GridlineColor =10921638

                    LayoutCachedLeft =9127
                    LayoutCachedTop =6242
                    LayoutCachedWidth =10374
                    LayoutCachedHeight =7185
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =9247
                            Top =6122
                            Width =510
                            Height =315
                            BorderColor =8355711
                            ForeColor =-2147483617
                            Name ="Bijschrift11"
                            Caption ="Naar"
                            GridlineColor =10921638
                            LayoutCachedLeft =9247
                            LayoutCachedTop =6122
                            LayoutCachedWidth =9757
                            LayoutCachedHeight =6437
                            BackThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =9313
                            Top =6480
                            OptionValue =1
                            BorderColor =10921638
                            Name ="Keuzerondje13"
                            GridlineColor =10921638

                            LayoutCachedLeft =9313
                            LayoutCachedTop =6480
                            LayoutCachedWidth =9573
                            LayoutCachedHeight =6720
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =9543
                                    Top =6452
                                    Width =555
                                    Height =315
                                    BorderColor =8355711
                                    ForeColor =-2147483617
                                    Name ="Bijschrift14"
                                    Caption ="Excel"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =9543
                                    LayoutCachedTop =6452
                                    LayoutCachedWidth =10098
                                    LayoutCachedHeight =6767
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =9313
                            Top =6810
                            TabIndex =1
                            OptionValue =2
                            BorderColor =10921638
                            Name ="Keuzerondje15"
                            GridlineColor =10921638

                            LayoutCachedLeft =9313
                            LayoutCachedTop =6810
                            LayoutCachedWidth =9573
                            LayoutCachedHeight =7050
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =9543
                                    Top =6782
                                    Width =645
                                    Height =315
                                    BorderColor =8355711
                                    ForeColor =-2147483617
                                    Name ="Bijschrift16"
                                    Caption ="Intern"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =9543
                                    LayoutCachedTop =6782
                                    LayoutCachedWidth =10188
                                    LayoutCachedHeight =7097
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =5669
                    Top =6808
                    Width =2271
                    Height =568
                    TabIndex =2
                    ForeColor =4210752
                    Name ="btnSchemaNaarOpstelling"
                    Caption ="Schema --> Opstelling"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =5669
                    LayoutCachedTop =6808
                    LayoutCachedWidth =7940
                    LayoutCachedHeight =7376
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
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =1133
                    Top =6292
                    Width =2271
                    Height =568
                    TabIndex =3
                    ForeColor =4210752
                    Name ="Knop18"
                    Caption ="Nieuwe ronde indelen"
                    GridlineColor =10921638

                    LayoutCachedLeft =1133
                    LayoutCachedTop =6292
                    LayoutCachedWidth =3404
                    LayoutCachedHeight =6860
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
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6009
                    Top =5839
                    Width =1800
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSessie"
                    Caption ="Sessie 11"
                    GridlineColor =10921638
                    LayoutCachedLeft =6009
                    LayoutCachedTop =5839
                    LayoutCachedWidth =7809
                    LayoutCachedHeight =6154
                End
                Begin Label
                    OverlapFlags =85
                    Left =566
                    Top =5669
                    Width =4455
                    Height =585
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Bijschrift24"
                    Caption ="Alle acties hebben als basis het interne schema\015\012dat is gedownload van het"
                        " werkbestand"
                    GridlineColor =10921638
                    LayoutCachedLeft =566
                    LayoutCachedTop =5669
                    LayoutCachedWidth =5021
                    LayoutCachedHeight =6254
                End
                Begin CommandButton
                    OverlapFlags =93
                    TextFontCharSet =177
                    Left =11338
                    Top =6236
                    Width =576
                    Height =576
                    TabIndex =4
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

                    LayoutCachedLeft =11338
                    LayoutCachedTop =6236
                    LayoutCachedWidth =11914
                    LayoutCachedHeight =6812
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
                Begin Subform
                    OverlapFlags =85
                    Width =11850
                    Height =5445
                    TabIndex =5
                    BorderColor =10921638
                    Name ="subSchema"
                    SourceObject ="Form.frmSchema"
                    LinkChildFields ="ToernooiID"
                    LinkMasterFields ="ID"
                    GridlineColor =10921638

                    LayoutCachedWidth =11850
                    LayoutCachedHeight =5445
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


Private Sub btnSchemaNaarOpstelling_Click()
'opfrissing per sessie
Dim Naar, van, tot As Integer
Dim db As Database
Dim rs As Recordset
Dim MySheet As Worksheet
Dim StartBook As Workbook
Dim strWorkfile As String
Dim question As Integer
Dim intSessienr As Integer
Dim intTeamnr As Integer
Dim sessieaanwezig As Integer
Dim rijteller, beginTel As Integer
Dim TestExcel As Integer
Dim TeamTegenstanders() As Long
Dim i, j, K, thuis, uit As Integer
Dim sql As String
Set db = CurrentDb

'test of er een werkbestand is
strWorkfile = WORKFOLDER & WORKFILE

If Not fnExists(strWorkfile) Then
     MsgBox ("Er is nog geen excel bestand aangemaakt")
     Exit Sub
End If

ReDim TeamTegenstanders(AANTALTEAMS, WEDSTRIJDENPERSESSIE)

'bepaal welke rondes

'
intSessienr = Sessienr
sessieoffset = GespeeldeWedstrijden(intSessienr, lngToernooi)

Select Case Me.grpUitvoernaar
Case 1
    'test of opstelling excel is aanwezig
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Application.Visible = intExcelZichtbaar
    xlApp.Application.DisplayAlerts = False
    Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)
    Set MySheet = StartBook.Worksheets("Import_Opstelling")
    sessieaanwezig = False
     rijteller = 2
  Do While MySheet.Cells(rijteller, 1) <> ""
    If MySheet.Cells(rijteller, 1).Value = intSessienr Then
        beginTel = rijteller
        sessieaanwezig = True
        Exit Do
    End If
    rijteller = rijteller + 1
  Loop
   If Not sessieaanwezig Then
        question = MsgBox("Er is nog geen opstelling aanwezig in de excel file")
        Set MySheet = Nothing
        Set StartBook = Nothing
        xlApp.Application.DisplayAlerts = True
        xlApp.Application.Quit
        Set xlApp = Nothing
        Exit Sub
    End If




Case 2
'test of opstelling intern is aanwezig
Set rs = db.OpenRecordset("select * from tblOpstelling where [ToernooiID] = " & lngToernooi & " and [Sessie] = " & CInt(intSessienr))
If rs.BOF And rs.EOF Then
    MsgBox ("Er is  geen opstelling geladen of aanwezig")
    rs.Close
    db.Close
    Exit Sub
End If

End Select
 
 

' team  ,  zoek tegenstander wedstrij1, zoek tegenstander wedstrijd 2


'sessieoffset = GespeeldeWedstrijden(intSessienr, lngToernooi)
van = sessieoffset + 1
Naar = sessieoffset + WEDSTRIJDENPERSESSIE
sql = ""
sql = sql & "Select * from tblSchema where "
sql = sql & " ([ToernooiID] = " & lngToernooi
sql = sql & ") and "
sql = sql & " ([Wedstrijdronde] Between " & van & " and " & Naar
sql = sql & ") Order By [Wedstrijdronde],[Paring];"


Set rs = db.OpenRecordset(sql)
rs.MoveFirst
i = rs!Wedstrijdronde
j = 1
Do While Not rs.EOF
    thuis = rs!TeamThuis
    uit = rs!TeamUit
    TeamTegenstanders(thuis, j) = uit
    TeamTegenstanders(uit, j) = thuis
    rs.MoveNext
    If rs.EOF Then Exit Do
    If i <> rs!Wedstrijdronde Then
        i = rs!Wedstrijdronde
        j = j + 1
    End If
Loop
rs.Close


Select Case Me.grpUitvoernaar
Case 1
    For i = 1 To AANTALTEAMS
        For j = 1 To WEDSTRIJDENPERSESSIE
            MySheet.Cells(rijteller, 6 + j).Value = TeamTegenstanders(i, j)
        Next
      rijteller = rijteller + 1
    Next
    StartBook.Save
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
Case 2
    Set rs = db.OpenRecordset("select * from tblOpstelling where [ToernooiID] = " & lngToernooi & " and [Sessie] = " & CInt(intSessienr))
    rs.MoveFirst
    Do While Not rs.EOF
        rs.Edit
        For j = 1 To WEDSTRIJDENPERSESSIE
            rs.Fields("Wedstrijd" & j) = TeamTegenstanders(rs!Teamnr, j)
        Next
        rs.Update
        rs.MoveNext
    Loop
    rs.Close
    db.Close
End Select

End Sub

Private Sub btnSchemaNaarUitslag_Click()
Select Case Me.grpUitvoernaar
Case 1
Call AddExcelTeamUitslagen_Schema(lngToernooi, lngSessie)
Case 2
Call AddInternTeamUitslagen_Schema(lngToernooi, lngSessie)
End Select

End Sub

Private Sub btnSluiten_Click()
    If CurrentProject.AllForms("Start_VT").IsLoaded = False Then
        DoCmd.Close
    Else
        DoCmd.BrowseTo acBrowseToForm, "frmBegin"
   End If
End Sub

Private Sub cboKiesRonde_AfterUpdate()
' filter subformulier
Dim sCriteria As String
 If Me.cboKiesRonde <> "Alle" Then
    sCriteria = "[WedstrijdRonde] = " & Me.cboKiesRonde
    [subSchema].Form.Filter = sCriteria
   
    [subSchema].Form.FilterOn = True
 Else
    [subSchema].Form.Filter = ""
    [subSchema].Form.FilterOn = False
 End If

End Sub



Private Sub Form_Open(Cancel As Integer)
    Me.lblSessie.Caption = "Sessie " & Sessienr
End Sub
