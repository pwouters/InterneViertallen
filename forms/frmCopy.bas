Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9188
    DatasheetFontHeight =11
    ItemSuffix =42
    Right =24675
    Bottom =12240
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0x7bf2d78ef590e540
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
        Begin EmptyCell
            Height =240
            GridlineColor =12632256
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =3741
            Name ="Details"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2205
                    Top =795
                    Width =6870
                    Height =735
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtBronbestand"
                    OnMouseDown ="[Event Procedure]"
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2205
                    LayoutCachedTop =795
                    LayoutCachedWidth =9075
                    LayoutCachedHeight =1530
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Top =795
                            Width =2148
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblBronbestand"
                            Caption ="Kies Model bestand"
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedTop =795
                            LayoutCachedWidth =2148
                            LayoutCachedHeight =1125
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2205
                    Top =1587
                    Width =6870
                    Height =675
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtDestinationFolder"
                    OnMouseDown ="[Event Procedure]"
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2205
                    LayoutCachedTop =1587
                    LayoutCachedWidth =9075
                    LayoutCachedHeight =2262
                    RowStart =1
                    RowEnd =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Top =1587
                            Width =2148
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblDestination"
                            Caption ="Kies Werkfolder"
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedTop =1587
                            LayoutCachedWidth =2148
                            LayoutCachedHeight =1917
                            RowStart =1
                            RowEnd =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2258
                    Top =2437
                    Width =6930
                    Height =315
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtNieuwExcel"
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2258
                    LayoutCachedTop =2437
                    LayoutCachedWidth =9188
                    LayoutCachedHeight =2752
                    RowStart =2
                    RowEnd =2
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Top =2437
                            Width =2156
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblNieuwExcel"
                            Caption ="Nieuwe Naam Excel"
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedTop =2437
                            LayoutCachedWidth =2156
                            LayoutCachedHeight =2752
                            RowStart =2
                            RowEnd =2
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3968
                    Top =2834
                    Width =2250
                    Height =658
                    TabIndex =3
                    ForeColor =4210752
                    Name ="btnCopy"
                    Caption ="Copieer Bestand"
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =3968
                    LayoutCachedTop =2834
                    LayoutCachedWidth =6218
                    LayoutCachedHeight =3492
                    RowStart =3
                    RowEnd =3
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =9
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =177
                    Left =6803
                    Top =2834
                    Width =576
                    Height =576
                    TabIndex =4
                    ForeColor =-2147483630
                    Name ="btnSluiten"
                    Caption ="Knop84"
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

                    LayoutCachedLeft =6803
                    LayoutCachedTop =2834
                    LayoutCachedWidth =7379
                    LayoutCachedHeight =3410
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

Private Sub txtBronbestand_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If right-button clicked


    If strTemplate_Folder = "" Then
        Dim db As Database
        Dim rs As Recordset
        'kies basis
        Set db = CurrentDb
        Set rs = db.OpenRecordset("tblFolders")
        strTemplate_Folder = rs!TemplateFolder
        rs.Close
        db.Close
        
    End If
    
    
    If Button = 1 Then
            Dim strFile As String
            strFile = GetFileName(strTemplate_Folder)
            If Not strFile = "" Then
                'strip path
                Me.txtBronbestand = strFile
            End If
    End If




End Sub

Private Sub txtDestinationFolder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    If strExcel_Folder = "" Then
         Dim db As Database
         Dim rs As Recordset
        'kies basis
        Set db = CurrentDb
        Set rs = db.OpenRecordset("tblFolders")
        strTemplate_Folder = rs!ExcelFolder
        rs.Close
        db.Close
    End If
    If Button = 1 Then
            Dim strFolder As String
            strFolder = GetFolderName(strExcel_Folder)
            If Not strFolder = "" Then
                'strip path
                Me.txtDestinationFolder = Left(strFolder, Len(strFolder) - 1)
            End If
    End If
End Sub
