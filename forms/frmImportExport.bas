Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =12132
    DatasheetFontHeight =11
    ItemSuffix =7
    Right =19140
    Bottom =12240
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0xee32c2fe4f92e540
    End
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
                    Caption ="Import / Export"
                    GridlineColor =10921638
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =969
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3401
                    Top =226
                    Width =8360
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblImport"
                    Caption =" Werkbestand = Interne_Viertallen.xlsx"
                    GridlineColor =10921638
                    LayoutCachedLeft =3401
                    LayoutCachedTop =226
                    LayoutCachedWidth =11761
                    LayoutCachedHeight =541
                End
            End
        End
        Begin Section
            Height =4195
            Name ="Details"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    Left =566
                    Top =963
                    Width =2091
                    Height =568
                    ForeColor =4210752
                    Name ="btnImportTeams"
                    Caption ="Import Teams"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =963
                    LayoutCachedWidth =2657
                    LayoutCachedHeight =1531
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
                    OverlapFlags =95
                    Left =566
                    Top =1530
                    Width =2091
                    Height =568
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnImportOpstelling"
                    Caption ="Import Opstelling"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =1530
                    LayoutCachedWidth =2657
                    LayoutCachedHeight =2098
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
                    Top =2097
                    Width =2091
                    Height =568
                    TabIndex =2
                    ForeColor =4210752
                    Name ="btnImportIndeling"
                    Caption ="Import Uitslagen"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =2097
                    LayoutCachedWidth =2657
                    LayoutCachedHeight =2665
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
                    OverlapFlags =93
                    Left =3401
                    Top =963
                    Width =2091
                    Height =568
                    TabIndex =3
                    ForeColor =4210752
                    Name ="btnExportTeams"
                    Caption ="Export Teams"
                    GridlineColor =10921638

                    LayoutCachedLeft =3401
                    LayoutCachedTop =963
                    LayoutCachedWidth =5492
                    LayoutCachedHeight =1531
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
                    OverlapFlags =95
                    Left =3401
                    Top =1530
                    Width =2091
                    Height =568
                    TabIndex =4
                    ForeColor =4210752
                    Name ="btnExportOpstelling"
                    Caption ="Export Opstelling"
                    GridlineColor =10921638

                    LayoutCachedLeft =3401
                    LayoutCachedTop =1530
                    LayoutCachedWidth =5492
                    LayoutCachedHeight =2098
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
                    OverlapFlags =87
                    Left =3401
                    Top =2097
                    Width =2091
                    Height =568
                    TabIndex =5
                    ForeColor =4210752
                    Name ="btnExportIndeling"
                    Caption ="Export Indeling"
                    GridlineColor =10921638

                    LayoutCachedLeft =3401
                    LayoutCachedTop =2097
                    LayoutCachedWidth =5492
                    LayoutCachedHeight =2665
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
                    TextFontCharSet =177
                    Left =4535
                    Top =3231
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

                    LayoutCachedLeft =4535
                    LayoutCachedTop =3231
                    LayoutCachedWidth =5111
                    LayoutCachedHeight =3807
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
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3401
                    Top =113
                    Width =4536
                    Height =315
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =3484194
                    Name ="cboKiesTabblad"
                    RowSourceType ="Value List"
                    RowSource ="Team_template;WebInfo;Teams;Schema;SchemaZwitsers;Kruistabel;TeamUitslagen;Impor"
                        "t_Opstelling;Import_Uitslag;VPSchaal;Imptabel;Avond_1_Teamnr_4;Avond_1_Teamnr_5;"
                        "Avond_1_Teamnr_6;Avond_1_Teamnr_7;Avond_1_Teamnr_8;Avond_1_Teamnr_9;Avond_1_Team"
                        "nr_10;Avond_1_Teamnr_11;Avond_1_Teamnr_12;Avond_1_Teamnr_14;Avond_1_Teamnr_15;Av"
                        "ond_2_Teamnr_2;Avond_2_Teamnr_3;Avond_2_Teamnr_5;Avond_2_Teamnr_7;Avond_2_Teamnr"
                        "_8;Avond_2_Teamnr_10;Avond_2_Teamnr_11;Avond_2_Teamnr_12;Avond_2_Teamnr_13;Avond"
                        "_2_Teamnr_14;Avond_2_Teamnr_15;Avond_2_Teamnr_6;Avond_2_Teamnr_9;Avond_1_Teamnr_"
                        "13;Avond_3_Teamnr_1;Avond_3_Teamnr_2;Avond_3_Teamnr_3;Avond_3_Teamnr_5;Avond_3_T"
                        "eamnr_6;Avond_3_Teamnr_7;Avond_3_Teamnr_8;Avond_3_Teamnr_9;Avond_3_Teamnr_10;Avo"
                        "nd_3_Teamnr_11;Avond_3_Teamnr_12;Avond_3_Teamnr_13;Avond_3_Teamnr_14;Avond_3_Tea"
                        "mnr_15;Avond_2_Teamnr_4;Avond_3_Teamnr_4;Avond_4_Teamnr_2;Avond_4_Teamnr_3;Avond"
                        "_4_Teamnr_4;Avond_4_Teamnr_5;Avond_4_Teamnr_6;Avond_4_Teamnr_7;Avond_4_Teamnr_8;"
                        "Avond_4_Teamnr_9;Avond_4_Teamnr_10;Avond_4_Teamnr_11;Avond_4_Teamnr_12;Avond_4_T"
                        "eamnr_13;Avond_4_Teamnr_14;Avond_4_Teamnr_15;Avond_4_Teamnr_1;Avond_2_Teamnr_1;A"
                        "vond_1_Teamnr_1;Avond_1_Teamnr_2;Avond_1_Teamnr_3"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3401
                    LayoutCachedTop =113
                    LayoutCachedWidth =7937
                    LayoutCachedHeight =428
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =566
                            Top =113
                            Width =2655
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Bijschrift2"
                            Caption ="Kies tabblad om te bekijken"
                            GridlineColor =10921638
                            LayoutCachedLeft =566
                            LayoutCachedTop =113
                            LayoutCachedWidth =3221
                            LayoutCachedHeight =428
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =6236
                    Top =1530
                    Width =2085
                    Height =570
                    TabIndex =8
                    ForeColor =4210752
                    Name ="btnInternOpstelling"
                    Caption ="Opstelling Intern"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6236
                    LayoutCachedTop =1530
                    LayoutCachedWidth =8321
                    LayoutCachedHeight =2100
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
                    Left =6236
                    Top =2097
                    Width =2085
                    Height =570
                    TabIndex =9
                    ForeColor =4210752
                    Name ="btnInternUitslagen"
                    Caption ="Uitslagen Intern"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6236
                    LayoutCachedTop =2097
                    LayoutCachedWidth =8321
                    LayoutCachedHeight =2667
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
                    Left =6236
                    Top =963
                    Width =2085
                    Height =570
                    TabIndex =10
                    ForeColor =4210752
                    Name ="btnInternTeams"
                    Caption ="Teams Intern"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6236
                    LayoutCachedTop =963
                    LayoutCachedWidth =8321
                    LayoutCachedHeight =1533
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

Private Sub btnImportIndeling_Click()
Dim db As Database
Dim rs As Recordset
Dim MySheet As Worksheet
Dim StartBook As Workbook
Dim strWorkFile As String
Dim question As Integer
Dim TeamsID() As Long
ReDim TeamsID(AANTALTEAMS)

Set db = CurrentDb
Set rs = db.OpenRecordset("select * from tblTeams where [ToernooiID] = " & lngToernooi)
If rs.BOF And rs.EOF Then
    MsgBox ("er zijn nog geen teams geimporteerd van dit toernooi")
    rs.Close
    db.Close
    Exit Sub
End If

rs.MoveFirst
Do While Not rs.EOF
    TeamsID(rs!Teamnr) = rs!Id
    rs.MoveNext
Loop
rs.Close


Set rs = db.OpenRecordset("select * from tblUitslagen where [ToernooiID] = " & lngToernooi & " and SessieID = " & lngSessie)

If Not (rs.BOF And rs.EOF) Then
  MsgBox ("Indeling is al geimporteerd")
    rs.Close
    db.Close
    Exit Sub
End If


'test of er een werkbestand is
strWorkFile = WORKFOLDER & WORKFILE

If Not fnExists(strWorkFile) Then
     MsgBox ("Er is nog geen excel bestand aangemaakt")
     Exit Sub
End If

     
Set xlApp = CreateObject("Excel.Application")
xlApp.Application.Visible = intExcelZichtbaar
xlApp.Application.DisplayAlerts = False
Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)
Set MySheet = StartBook.Worksheets("Teamuitslagen")

If MySheet.Cells(2, 1) = "" Then
    question = MsgBox("Er zijn geen teams aanwezig")
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    Exit Sub
End If
    
 rijteller = 2
 
 Do While MySheet.Cells(rijteller, 1) <> "" And MySheet.Cells(rijteller, 1) <> Sessienr
    rijteller = rijteller + 1
 Loop
    If MySheet.Cells(rijteller, 1) = "" Then
       MsgBox ("De indeling is nog niet ingevoerd")
        Exit Sub
    End If
    
Set rs = db.OpenRecordset("select * from tblUitslagen where [ToernooiID] = " & lngToernooi & " and SessieID = " & lngSessie)
If Not (rs.BOF And rs.EOF) Then
    MsgBox ("De indeling is reeds geimporteerd")
    'Controleren of er verschillen zijn
    'Verwijderen en opnieuw laden
    
    
    Exit Sub
End If


Set rs = db.OpenRecordset("tblUitslagen")

 Do While MySheet.Cells(rijteller, 1) = Sessienr
    rs.AddNew
    rs!ToernooiID = lngToernooi
    rs!SessieID = lngSessie
    rs!Wedstrijdnr = MySheet.Cells(rijteller, 2).Value
    rs!TeamIDThuis = TeamsID(MySheet.Cells(rijteller, 3).Value)
    rs!TeamIDUit = TeamsID(MySheet.Cells(rijteller, 4).Value)
    If MySheet.Cells(rijteller, 7).Value <> "" Then
        rs!ImpsThuis = MySheet.Cells(rijteller, 7).Value
    End If
    If MySheet.Cells(rijteller, 8).Value <> "" Then
        rs!ImpsUit = MySheet.Cells(rijteller, 8).Value
    End If
    If MySheet.Cells(rijteller, 9).Value <> "" Then
        rs!VpsThuis = MySheet.Cells(rijteller, 9).Value
    End If
    If MySheet.Cells(rijteller, 10).Value <> "" Then
        rs!VpsUit = MySheet.Cells(rijteller, 10).Value
    End If
    rs.Update
    rijteller = rijteller + 1
 Loop
    
 rs.Close
 db.Close
 
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
 
End Sub


Private Sub btnImportOpstelling_Click()
Dim rijteller As Integer
Dim db As Database
Dim rs As Recordset
Dim MySheet As Worksheet
Dim StartBook As Workbook
Dim strWorkFile As String
Dim question As Integer
Dim intSessienr As Integer
Dim intTeamnr As Integer
Dim sessieaanwezig As Integer
Dim TestExcel As Integer
Dim TeamsID() As Long
ReDim TeamsID(AANTALTEAMS)

Set db = CurrentDb
Set rs = db.OpenRecordset("select * from tblTeams where [ToernooiID] = " & lngToernooi)
If rs.BOF And rs.EOF Then
    MsgBox ("er zijn nog geen teams geimporteerd van dit toernooi")
    rs.Close
    db.Close
    Exit Sub
End If

rs.MoveFirst
Do While Not rs.EOF
    TeamsID(rs!Teamnr) = rs!Id
    rs.MoveNext
Loop
rs.Close

intSessienr = Sessienr


Set rs = db.OpenRecordset("select * from tblOpstelling where [ToernooiID] = " & lngToernooi & " and [Sessie] = " & CInt(intSessienr))
If Not (rs.BOF And rs.EOF) Then
    MsgBox ("Er is  reeds opstelling geladen of aanwezig")
    rs.Close
    db.Close
    Exit Sub
End If

'test of er een werkbestand is
strWorkFile = WORKFOLDER & WORKFILE

If Not fnExists(strWorkFile) Then
     MsgBox ("Er is nog geen excel bestand aangemaakt")
     Exit Sub
End If

     
Set xlApp = CreateObject("Excel.Application")
xlApp.Application.Visible = intExcelZichtbaar
xlApp.Application.DisplayAlerts = False
Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)
Set MySheet = StartBook.Worksheets("Import_Opstelling")

If MySheet.Cells(2, 1) = "" Then
    question = MsgBox("Er zijn geen opstelling aanwezig")
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    Exit Sub
End If
    sessieaanwezig = False
 rijteller = 2
  Do While MySheet.Cells(rijteller, 1) <> ""
    If MySheet.Cells(rijteller, 1).Value = intSessienr Then
            If Not sessieaanwezig Then sessieaanwezig = True
            rs.AddNew
            rs!ToernooiID = lngToernooi
            rs!SessieID = lngSessie
            rs!Sessie = MySheet.Cells(rijteller, 1).Value
            intTeamnr = MySheet.Cells(rijteller, 2).Value
            rs!Teamnr = intTeamnr
            rs!TeamID = TeamsID(intTeamnr)
            rs!Speler1 = MySheet.Cells(rijteller, 3).Value
            rs!Speler2 = MySheet.Cells(rijteller, 4).Value
            rs!Speler3 = MySheet.Cells(rijteller, 5).Value
            rs!Speler4 = MySheet.Cells(rijteller, 6).Value
            Kolom = 7
            intWedstrijd = 1
             Do While MySheet.Cells(rijteller, Kolom).Value <> "" And intWedstrijd < 12
                rs.Fields("wedstrijd" & intWedstrijd) = MySheet.Cells(rijteller, Kolom).Value
                Kolom = Kolom + 1
                intWedstrijd = intWedstrijd + 1
             Loop
            rs.Update
     End If
    rijteller = rijteller + 1
   Loop
    
    rs.Close
    db.Close
  
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    If Not sessieaanwezig Then
        question = MsgBox("Er is geen opstelling aanwezig in de excel file")
    End If
 End Sub

Private Sub btnImportTeams_Click()
  'kijk eerst op er al teams opgenomen zijn
Dim xlApp As Object
Dim TestExcel As Integer

Dim db As Database
Dim rs As Recordset
Dim MySheet As Worksheet
Dim StartBook As Workbook
Dim strWorkFile As String
Dim question As Integer
Set db = CurrentDb
Set rs = db.OpenRecordset("select * from tblTeams where [ToernooiID] = " & lngToernooi)

If Not (rs.BOF And rs.EOF) Then
    MsgBox ("Er zijn reeds teams geladen of aanwezig")
    rs.Close
    db.Close
    
    Exit Sub
End If

'test of er een werkbestand is
strWorkFile = WORKFOLDER & WORKFILE

If Not fnExists(strWorkFile) Then
     MsgBox ("Er is nog geen excel bestand aangemaakt")
     Exit Sub
End If

     
Set xlApp = CreateObject("Excel.Application")
xlApp.Application.Visible = intExcelZichtbaar
xlApp.Application.DisplayAlerts = False
Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)
Set MySheet = StartBook.Worksheets("Teams")

If MySheet.Cells(2, 1) = "" Then
    question = MsgBox("Er zijn geen teams aanwezig")
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    Exit Sub
End If
    
 rijteller = 2
 TestExcel = False
 Do While MySheet.Cells(rijteller, 1) <> ""
    rs.AddNew
    rs!ToernooiID = lngToernooi
    rs!Teamnr = MySheet.Cells(rijteller, 1).Value
    rs!TeamNaam = MySheet.Cells(rijteller, 2).Value
    'niet meer dan 8 spelers
    Kolom = 3
    speler = 1
    Do While MySheet.Cells(rijteller, Kolom).Value <> "" And speler < 9
       rs.Fields("speler" & speler) = MySheet.Cells(rijteller, Kolom).Value
       Kolom = Kolom + 1
       speler = speler + 1
    Loop
    rs.Update
    rijteller = rijteller + 1
 Loop
 rs.Close
 db.Close
 
    Set MySheet = Nothing
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
 
End Sub

Private Sub btnInternOpstelling_Click()
    DoCmd.OpenForm "frmOpstelling", acFormDS
End Sub

Private Sub btnInternTeams_Click()
    DoCmd.OpenForm "frmTeams", acFormDS
End Sub

Private Sub btnInternUitslagen_Click()
DoCmd.OpenForm "frmTeamUitslagen", acFormDS
End Sub

Private Sub btnSluiten_Click()
     If CurrentProject.AllForms("Start_VT").IsLoaded = False Then
        DoCmd.Close
    Else
        DoCmd.BrowseTo acBrowseToForm, "frmBegin"
   End If
End Sub

Private Sub cboKiesTabblad_AfterUpdate()
Dim xlApp As Object
Dim sheetname As String
sheetname = Trim(Me.cboKiesTabblad)
Me.cboKiesTabblad = ""
Dim MySheet As Worksheet
Dim StartBook As Workbook

Set xlApp = CreateObject("Excel.Application")

xlApp.Application.Visible = True
xlApp.Application.DisplayAlerts = False

Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)
Set MySheet = StartBook.Worksheets(sheetname)
'test workfile
    'Me.cboKiesTabblad.Clear
 MySheet.Activate
 
    
    'Set StartBook = Nothing
  '  xlApp.Application.DisplayAlerts = True
   ' xlApp.Application.Quit
    'Set xlApp = Nothing
End Sub

Private Sub Form_Open(Cancel As Integer)
' vulcombo box
Dim StartBook As Object
Dim MySheets As Object
Dim lijst As String
Dim fCount As Integer
If Not fnExists(WORKFOLDER & WORKFILE) Then
    MsgBox ("applicatie kan " & WORKFOLDER & WORKFILE & "niet vinden ")
    Exit Sub
End If


Me.lblImport.Caption = " Werkbestand = " & WORKFILE
'test op opstelling


Set xlApp = CreateObject("Excel.Application")

xlApp.Application.Visible = False
xlApp.Application.DisplayAlerts = False

Set StartBook = xlApp.Workbooks.Open(WORKFOLDER & WORKFILE)

'test workfile
    'Me.cboKiesTabblad.Clear
    
    Dim strWs As String
    Dim i As Integer
    fCount = StartBook.Sheets.Count
    
     SysCmd acSysCmdInitMeter, "laad tabblad namen van het werkbestand...", fCount
 
    For i = 1 To fCount
    SysCmd acSysCmdUpdateMeter, fCount
        lijst = lijst & StartBook.Sheets(i).name
        If i <> StartBook.Sheets.Count Then
         lijst = lijst & ";"
        End If
    Next
 
    Me.cboKiesTabblad.RowSource = lijst
     
    Set StartBook = Nothing
    xlApp.Application.DisplayAlerts = True
    xlApp.Application.Quit
    Set xlApp = Nothing
    SysCmd acSysCmdRemoveMeter
End Sub
