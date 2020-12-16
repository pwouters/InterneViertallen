Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10209
    DatasheetFontHeight =11
    ItemSuffix =39
    Left =2520
    Top =1200
    Right =17445
    Bottom =11070
    DatasheetGridlinesColor =15132391
    OrderBy ="[qryWedstrijden_Kruistabel].[Totaal VPS] DESC, [qryWedstrijden_Kruistabel].[Team"
        "nr]"
    RecSrcDt = Begin
        0x00211c2f1392e540
    End
    RecordSource ="qryWedstrijden_Kruistabel"
    Caption ="Stand"
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
        Begin FormHeader
            Height =1026
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
                    Left =57
                    Top =57
                    Width =5328
                    Height =969
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Bijschrift24"
                    Caption ="Stand "
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =5385
                    LayoutCachedHeight =1026
                End
            End
        End
        Begin Section
            Height =8900
            Name ="Details"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =342
                    Width =1050
                    Height =330
                    ColumnWidth =1050
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Teamnr"
                    ControlSource ="Teamnr"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =342
                    LayoutCachedWidth =3942
                    LayoutCachedHeight =672
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =342
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Teamnr_Bijschrift"
                            Caption ="Teamnr"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =342
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =672
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =741
                    Width =7260
                    Height =600
                    ColumnWidth =2925
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="TeamNaam"
                    ControlSource ="TeamNaam"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =741
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =1341
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =741
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="TeamNaam_Bijschrift"
                            Caption ="TeamNaam"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =741
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =1071
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =1425
                    Width =3660
                    Height =330
                    ColumnWidth =1380
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Totaal VPS"
                    ControlSource ="Totaal VPS"
                    Format ="#,#00"
                    EventProcPrefix ="Totaal_VPS"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =1425
                    LayoutCachedWidth =6552
                    LayoutCachedHeight =1755
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =1425
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Totaal VPS_Bijschrift"
                            Caption ="Totaal VPS"
                            EventProcPrefix ="Totaal_VPS_Bijschrift"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =1425
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =1755
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =1824
                    Width =3660
                    Height =330
                    ColumnWidth =1020
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="GemVanVPS"
                    ControlSource ="GemVanVPS"
                    Format ="#,#00"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =1824
                    LayoutCachedWidth =6552
                    LayoutCachedHeight =2154
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =1824
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="GemVanVPS_Bijschrift"
                            Caption ="GemVanVPS"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =1824
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =2154
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =2223
                    Width =3660
                    Height =330
                    ColumnWidth =885
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="1_txt"
                    ControlSource ="1"
                    Format ="Standard"
                    EventProcPrefix ="Ctl1_txt"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =2223
                    LayoutCachedWidth =6552
                    LayoutCachedHeight =2553
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =2223
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="1_Bijschrift"
                            Caption ="1"
                            EventProcPrefix ="Ctl1_Bijschrift"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =2223
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =2553
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =2622
                    Width =3660
                    Height =330
                    ColumnWidth =915
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="2_txt"
                    ControlSource ="2"
                    Format ="Standard"
                    EventProcPrefix ="Ctl2_txt"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =2622
                    LayoutCachedWidth =6552
                    LayoutCachedHeight =2952
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =2622
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="2_Bijschrift"
                            Caption ="2"
                            EventProcPrefix ="Ctl2_Bijschrift"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =2622
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =2952
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =3021
                    Width =3660
                    Height =330
                    ColumnWidth =885
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="3_txt"
                    ControlSource ="3"
                    Format ="Standard"
                    EventProcPrefix ="Ctl3_txt"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =3021
                    LayoutCachedWidth =6552
                    LayoutCachedHeight =3351
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =3021
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="3_Bijschrift"
                            Caption ="3"
                            EventProcPrefix ="Ctl3_Bijschrift"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =3021
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =3351
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =3420
                    Width =3660
                    Height =330
                    ColumnWidth =915
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="4_txt"
                    ControlSource ="4"
                    Format ="Standard"
                    EventProcPrefix ="Ctl4_txt"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =3420
                    LayoutCachedWidth =6552
                    LayoutCachedHeight =3750
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =3420
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="4_Bijschrift"
                            Caption ="4"
                            EventProcPrefix ="Ctl4_Bijschrift"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =3420
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =3750
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =3819
                    Width =3660
                    Height =330
                    ColumnWidth =870
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="5_txt"
                    ControlSource ="5"
                    Format ="Standard"
                    EventProcPrefix ="Ctl5_txt"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =3819
                    LayoutCachedWidth =6552
                    LayoutCachedHeight =4149
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =3819
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="5_Bijschrift"
                            Caption ="5"
                            EventProcPrefix ="Ctl5_Bijschrift"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =3819
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =4149
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =4218
                    Width =3660
                    Height =330
                    ColumnWidth =900
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="6_txt"
                    ControlSource ="6"
                    Format ="Standard"
                    EventProcPrefix ="Ctl6_txt"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =4218
                    LayoutCachedWidth =6552
                    LayoutCachedHeight =4548
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =4218
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="6_Bijschrift"
                            Caption ="6"
                            EventProcPrefix ="Ctl6_Bijschrift"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =4218
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =4548
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =4617
                    Width =3660
                    Height =330
                    ColumnWidth =990
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="7_txt"
                    ControlSource ="7"
                    Format ="Standard"
                    EventProcPrefix ="Ctl7_txt"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =4617
                    LayoutCachedWidth =6552
                    LayoutCachedHeight =4947
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =4617
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="7_Bijschrift"
                            Caption ="7"
                            EventProcPrefix ="Ctl7_Bijschrift"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =4617
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =4947
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =5016
                    Width =3660
                    Height =330
                    ColumnWidth =975
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="8_txt"
                    ControlSource ="8"
                    Format ="Standard"
                    EventProcPrefix ="Ctl8_txt"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =5016
                    LayoutCachedWidth =6552
                    LayoutCachedHeight =5346
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =5016
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="8_Bijschrift"
                            Caption ="8"
                            EventProcPrefix ="Ctl8_Bijschrift"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =5016
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =5346
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =5466
                    Width =3660
                    Height =330
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="9_txt"
                    Format ="Standard"
                    EventProcPrefix ="Ctl9_txt"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =5466
                    LayoutCachedWidth =6552
                    LayoutCachedHeight =5796
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =342
                            Top =5466
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="9_Bijschrift"
                            Caption ="9"
                            EventProcPrefix ="Ctl9_Bijschrift"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =5466
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =5796
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2946
                    Top =5952
                    Width =3660
                    Height =330
                    TabIndex =13
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="10_txt"
                    Format ="Standard"
                    EventProcPrefix ="Ctl10_txt"
                    GridlineColor =10921638

                    LayoutCachedLeft =2946
                    LayoutCachedTop =5952
                    LayoutCachedWidth =6606
                    LayoutCachedHeight =6282
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =336
                            Top =5952
                            Width =2520
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="10_Bijschrift"
                            Caption ="10"
                            EventProcPrefix ="Ctl10_Bijschrift"
                            GridlineColor =10921638
                            LayoutCachedLeft =336
                            LayoutCachedTop =5952
                            LayoutCachedWidth =2856
                            LayoutCachedHeight =6282
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =6402
                    Width =3660
                    Height =330
                    TabIndex =14
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="11_txt"
                    Format ="Standard"
                    EventProcPrefix ="Ctl11_txt"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =6402
                    LayoutCachedWidth =6552
                    LayoutCachedHeight =6732
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =342
                            Top =6402
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="11_Bijschrijft"
                            Caption ="11"
                            EventProcPrefix ="Ctl11_Bijschrijft"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =6402
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =6732
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2946
                    Top =6888
                    Width =3660
                    Height =330
                    TabIndex =15
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="12_txt"
                    Format ="Standard"
                    EventProcPrefix ="Ctl12_txt"
                    GridlineColor =10921638

                    LayoutCachedLeft =2946
                    LayoutCachedTop =6888
                    LayoutCachedWidth =6606
                    LayoutCachedHeight =7218
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =396
                            Top =6888
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="12_Bijschrijft"
                            Caption ="12"
                            EventProcPrefix ="Ctl12_Bijschrijft"
                            GridlineColor =10921638
                            LayoutCachedLeft =396
                            LayoutCachedTop =6888
                            LayoutCachedWidth =2856
                            LayoutCachedHeight =7218
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =7338
                    Width =3660
                    Height =330
                    TabIndex =16
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="13_txt"
                    Format ="Standard"
                    EventProcPrefix ="Ctl13_txt"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =7338
                    LayoutCachedWidth =6552
                    LayoutCachedHeight =7668
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =342
                            Top =7338
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="13_Bijschrift"
                            Caption ="13"
                            EventProcPrefix ="Ctl13_Bijschrift"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =7338
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =7668
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2946
                    Top =7824
                    Width =3660
                    Height =330
                    TabIndex =17
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="14_txt"
                    Format ="Standard"
                    EventProcPrefix ="Ctl14_txt"
                    GridlineColor =10921638

                    LayoutCachedLeft =2946
                    LayoutCachedTop =7824
                    LayoutCachedWidth =6606
                    LayoutCachedHeight =8154
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =396
                            Top =7824
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="14_Bijschrift"
                            Caption ="14"
                            EventProcPrefix ="Ctl14_Bijschrift"
                            GridlineColor =10921638
                            LayoutCachedLeft =396
                            LayoutCachedTop =7824
                            LayoutCachedWidth =2856
                            LayoutCachedHeight =8154
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2946
                    Top =8274
                    Width =3660
                    Height =330
                    TabIndex =18
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="15_txt"
                    Format ="Standard"
                    EventProcPrefix ="Ctl15_txt"
                    GridlineColor =10921638

                    LayoutCachedLeft =2946
                    LayoutCachedTop =8274
                    LayoutCachedWidth =6606
                    LayoutCachedHeight =8604
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =396
                            Top =8274
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="15_Bijschrift"
                            Caption ="15"
                            EventProcPrefix ="Ctl15_Bijschrift"
                            GridlineColor =10921638
                            LayoutCachedLeft =396
                            LayoutCachedTop =8274
                            LayoutCachedWidth =2856
                            LayoutCachedHeight =8604
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

Private Sub Form_Open(Cancel As Integer)
Dim db As Database
Dim rs As Recordset
Dim fld As Field
Dim intKolommen, i As Integer

Set db = CurrentDb
Set rs = Me.Recordset
intKolommen = 0
For Each fld In rs.Fields
If Val(fld.name) > 0 Then
    intKolommen = intKolommen + 1
End If
Next
'nu verberg de kolommen

For i = intKolommen + 1 To 14
    Me.Controls(i & "_txt").ColumnHidden = True
Next




End Sub
