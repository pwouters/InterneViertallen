Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =12018
    DatasheetFontHeight =11
    ItemSuffix =24
    Left =2520
    Top =1200
    Right =17445
    Bottom =11070
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
                Begin Subform
                    OverlapFlags =85
                    Left =56
                    Width =11850
                    Height =5445
                    BorderColor =10921638
                    Name ="subSchema"
                    SourceObject ="Form.frmSchema"
                    LinkChildFields ="ToernooiID"
                    LinkMasterFields ="ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =56
                    LayoutCachedWidth =11906
                    LayoutCachedHeight =5445
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =5669
                    Top =6236
                    Width =2271
                    Height =568
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnSchemaNaarUitslag"
                    Caption ="Schema --> Uitslag"
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
                    TabIndex =2
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
                    TabIndex =3
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
                    OverlapFlags =85
                    Left =1133
                    Top =6292
                    Width =2271
                    Height =568
                    TabIndex =4
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
                    Caption ="Sessie"
                    GridlineColor =10921638
                    LayoutCachedLeft =6009
                    LayoutCachedTop =5839
                    LayoutCachedWidth =7809
                    LayoutCachedHeight =6154
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
Dim strWorkFile As String
Dim question As Integer
Dim intSessienr As Integer
Dim intTeamnr As Integer
Dim sessieaanwezig As Integer
Dim rijteller, beginTel As Integer
Dim TestExcel As Integer
Dim TeamTegenstanders() As Long
Dim i, j, k, thuis, uit As Integer
Dim sql As String
Set db = CurrentDb

'test of er een werkbestand is
strWorkFile = WORKFOLDER & WORKFILE

If Not fnExists(strWorkFile) Then
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
    Me.lblSessie.Caption = "Sessie " & lngSessie
End Sub
