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
    ItemSuffix =28
    Left =2520
    Top =1200
    Right =17190
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
            Height =8787
            Name ="Details"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    Left =6293
                    Top =6860
                    Width =2271
                    Height =568
                    ForeColor =4210752
                    Name ="btnSchemaNaarUitslag"
                    Caption ="Schema --> Uitslag"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6293
                    LayoutCachedTop =6860
                    LayoutCachedWidth =8564
                    LayoutCachedHeight =7428
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
                    Left =9184
                    Top =6866
                    Width =1247
                    Height =943
                    TabIndex =1
                    BorderColor =10921638
                    Name ="grpUitvoerNaar"
                    DefaultValue ="1"
                    GridlineColor =10921638

                    LayoutCachedLeft =9184
                    LayoutCachedTop =6866
                    LayoutCachedWidth =10431
                    LayoutCachedHeight =7809
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =9304
                            Top =6746
                            Width =510
                            Height =315
                            BorderColor =8355711
                            ForeColor =-2147483617
                            Name ="Bijschrift11"
                            Caption ="Naar"
                            GridlineColor =10921638
                            LayoutCachedLeft =9304
                            LayoutCachedTop =6746
                            LayoutCachedWidth =9814
                            LayoutCachedHeight =7061
                            BackThemeColorIndex =-1
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =9370
                            Top =7104
                            OptionValue =1
                            BorderColor =10921638
                            Name ="Keuzerondje13"
                            GridlineColor =10921638

                            LayoutCachedLeft =9370
                            LayoutCachedTop =7104
                            LayoutCachedWidth =9630
                            LayoutCachedHeight =7344
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =9600
                                    Top =7076
                                    Width =555
                                    Height =315
                                    BorderColor =8355711
                                    ForeColor =-2147483617
                                    Name ="Bijschrift14"
                                    Caption ="Excel"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =9600
                                    LayoutCachedTop =7076
                                    LayoutCachedWidth =10155
                                    LayoutCachedHeight =7391
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            Left =9370
                            Top =7434
                            TabIndex =1
                            OptionValue =2
                            BorderColor =10921638
                            Name ="Keuzerondje15"
                            GridlineColor =10921638

                            LayoutCachedLeft =9370
                            LayoutCachedTop =7434
                            LayoutCachedWidth =9630
                            LayoutCachedHeight =7674
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =9600
                                    Top =7406
                                    Width =645
                                    Height =315
                                    BorderColor =8355711
                                    ForeColor =-2147483617
                                    Name ="Bijschrift16"
                                    Caption ="Intern"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =9600
                                    LayoutCachedTop =7406
                                    LayoutCachedWidth =10245
                                    LayoutCachedHeight =7721
                                    ForeThemeColorIndex =-1
                                    ForeTint =100.0
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =6293
                    Top =7432
                    Width =2271
                    Height =568
                    TabIndex =2
                    ForeColor =4210752
                    Name ="btnSchemaNaarOpstelling"
                    Caption ="Schema --> Opstelling"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6293
                    LayoutCachedTop =7432
                    LayoutCachedWidth =8564
                    LayoutCachedHeight =8000
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
                    OverlapFlags =93
                    Left =566
                    Top =7370
                    Width =2271
                    Height =568
                    TabIndex =3
                    ForeColor =4210752
                    Name ="btnCreeerNieuweRonde"
                    Caption ="Wedstrijd indelen"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =7370
                    LayoutCachedWidth =2837
                    LayoutCachedHeight =7938
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
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6462
                    Top =6406
                    Width =1800
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSessie"
                    Caption ="Sessie 7"
                    GridlineColor =10921638
                    LayoutCachedLeft =6462
                    LayoutCachedTop =6406
                    LayoutCachedWidth =8262
                    LayoutCachedHeight =6721
                End
                Begin Label
                    OverlapFlags =85
                    Left =566
                    Top =5669
                    Width =4530
                    Height =1125
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Bijschrift24"
                    Caption ="Het geheel is een wisselwerking tussen excel werkbestand en interne bestanden, j"
                        "e kunt intern een nieuwe indeling maken en daarna verwerken in het excel werkbes"
                        "tand"
                    GridlineColor =10921638
                    LayoutCachedLeft =566
                    LayoutCachedTop =5669
                    LayoutCachedWidth =5096
                    LayoutCachedHeight =6794
                End
                Begin CommandButton
                    OverlapFlags =85
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
                Begin CommandButton
                    OverlapFlags =87
                    Left =566
                    Top =6803
                    Width =2271
                    Height =568
                    TabIndex =6
                    ForeColor =4210752
                    Name ="btnOpnieuwIndelen"
                    Caption ="Opnieuw Indelen"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =6803
                    LayoutCachedWidth =2837
                    LayoutCachedHeight =7371
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


Private Sub btnCreeerNieuweRonde_Click()
DoCmd.OpenForm "frmIndelingmaken", acNormal
End Sub

Private Sub btnOpnieuwIndelen_Click()
DoCmd.OpenForm "frmIndelingmaken", acNormal
End Sub

Private Sub btnSchemaNaarOpstelling_Click()
'opfrissing per sessie
Dim Naar, van, Tot As Integer
Dim db As Database
Dim rs As Recordset
Dim MySheet As Worksheet
Dim StartBook As Workbook
Dim strWorkfile As String
Dim Question As Integer
Dim intSessienr As Integer
Dim intTeamnr As Integer
Dim sessieaanwezig As Integer
Dim rijteller, beginTel As Integer
Dim TestExcel As Integer
Dim TeamTegenstanders() As Long
Dim i, j, k, Thuis, Uit As Integer
Dim sql As String
Set db = CurrentDb

'test of er een werkbestand is
strWorkfile = WORKFOLDER & WORKFILE

If Not fnExists(strWorkfile) Then
     MsgBox ("Er is nog geen excel bestand aangemaakt")
     Exit Sub
End If

ReDim TeamTegenstanders(AANTALTEAMS * 2, WEDSTRIJDENPERSESSIE)

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
        Question = MsgBox("Er is nog geen opstelling aanwezig in de excel file")
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
    Thuis = rs!TeamThuis
    Uit = rs!TeamUit
    TeamTegenstanders(Thuis, j) = Uit
    TeamTegenstanders(Uit, j) = Thuis
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
    intWedstrijdronde = Int(cboKiesRonde)
    sCriteria = "[WedstrijdRonde] = " & Me.cboKiesRonde
    [subSchema].Form.Filter = sCriteria
   
    [subSchema].Form.FilterOn = True
 Else
    
    [subSchema].Form.Filter = ""
    [subSchema].Form.FilterOn = False
 End If
 Call IndelenKnoppenZichtbaar
End Sub



Private Sub Form_Open(Cancel As Integer)
    Me.lblSessie.Caption = "Sessie " & Sessienr
    Me.cboKiesRonde = "Alle"
    
    Call IndelenKnoppenZichtbaar
End Sub



Private Sub IndelenKnoppenZichtbaar()
If Me.cboKiesRonde = "Alle" Then
    Me.btnOpnieuwIndelen.Visible = False
Else
    ' Uitslag is nog niet berekenend
    '
    'link tabel
    Call ImportTempTeamUitslagenTabel(WORKFOLDER & WORKFILE)
    Dim db As Database
    Dim rs As Recordset
    Dim sql As String
    Set db = CurrentDb
    sql = "Select * from tbl_temp_Uitslagen where [Wedstrijd] = " & Me.cboKiesRonde & " and Not isnull([ImpsThuis])"
    Set rs = db.OpenRecordset(sql)
    If rs.BOF And rs.EOF Then
        Me.btnOpnieuwIndelen.Visible = True
    Else
        Me.btnOpnieuwIndelen.Visible = False
    
    End If
    rs.Close
    db.Close
End If





End Sub
