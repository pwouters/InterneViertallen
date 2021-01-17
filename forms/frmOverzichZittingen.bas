Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6994
    DatasheetFontHeight =11
    ItemSuffix =1
    Right =13995
    Bottom =10470
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0x33ec3146fa91e540
    End
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
        Begin Section
            Height =5952
            Name ="Details"
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =1133
                    Top =623
                    Width =3105
                    Height =510
                    ForeColor =4210752
                    Name ="btnTeamKruisTabel"
                    Caption ="Overzicht wedstrijden per team"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1133
                    LayoutCachedTop =623
                    LayoutCachedWidth =4238
                    LayoutCachedHeight =1133
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

Private Sub btnTeamKruisTabel_Click()
Dim sql As String
sql = sql & "TRANSFORM Sum(qryWedstrijden.VPS) AS SomVanVPS "
sql = sql & "SELECT qryWedstrijden.Teamnr, qryWedstrijden.TeamNaam, Sum(qryWedstrijden.VPS) AS [Totaal VPS], Avg(qryWedstrijden.VPS) AS GemVanVPS "
sql = sql & "From qryWedstrijden "
sql = sql & "WHERE (((qryWedstrijden.ToernooiID) = " & lngToernooi & "))"
sql = sql & "GROUP BY qryWedstrijden.Teamnr, qryWedstrijden.TeamNaam "
sql = sql & "PIVOT qryWedstrijden.ZittingNr; "


End Sub
