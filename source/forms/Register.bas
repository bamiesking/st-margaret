Version =21
VersionRequired =20
Begin Form
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6994
    DatasheetFontHeight =11
    ItemSuffix =6
    Left =1613
    Top =3510
    Right =9150
    Bottom =9458
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0x440f8fcfd163e540
    End
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            Width =1701
            Height =1701
            BorderColor =16777215
            GridlineColor =16777215
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
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
        Begin Section
            Height =5952
            Name ="Detail"
            AutoHeight =1
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
                    Left =1364
                    Top =2097
                    Width =3964
                    Height =563
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="RegSelBox"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Activities].[ID], [Activities].[Activity Name] FROM Activities; "
                    ColumnWidths ="0;1440"
                    GridlineColor =10921638

                    LayoutCachedLeft =1364
                    LayoutCachedTop =2097
                    LayoutCachedWidth =5328
                    LayoutCachedHeight =2660
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1360
                            Top =1757
                            Width =1388
                            Height =320
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Activity Name_Label"
                            Caption ="Select activity"
                            EventProcPrefix ="Activity_Name_Label"
                            GridlineColor =10921638
                            LayoutCachedLeft =1360
                            LayoutCachedTop =1757
                            LayoutCachedWidth =2748
                            LayoutCachedHeight =2077
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1360
                    Top =3231
                    Height =567
                    TabIndex =1
                    ForeColor =4210752
                    Name ="Command4"
                    Caption ="Create Register"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1360
                    LayoutCachedTop =3231
                    LayoutCachedWidth =3061
                    LayoutCachedHeight =3798
                    BackColor =15123357
                    BorderColor =15123357
                    HoverColor =15652797
                    PressedColor =11957550
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =4
                    WebImagePaddingRight =3
                    WebImagePaddingBottom =3
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3629
                    Top =3234
                    Height =567
                    TabIndex =2
                    ForeColor =4210752
                    Name ="Command5"
                    Caption ="Return to Menu"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3629
                    LayoutCachedTop =3234
                    LayoutCachedWidth =5330
                    LayoutCachedHeight =3801
                    BackColor =15123357
                    BorderColor =15123357
                    HoverColor =15652797
                    PressedColor =11957550
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =4
                    WebImagePaddingRight =3
                    WebImagePaddingBottom =3
                End
                Begin Image
                    PictureType =2
                    Left =1644
                    Top =566
                    Width =3458
                    Height =1077
                    BorderColor =10921638
                    Name ="Image3"
                    Picture ="header-logo-retina"
                    GridlineColor =10921638

                    LayoutCachedLeft =1644
                    LayoutCachedTop =566
                    LayoutCachedWidth =5102
                    LayoutCachedHeight =1643
                    TabIndex =3
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

Private Sub Command4_Click()
    DoCmd.OpenReport "rptRegister", acViewReport, "", "", acNormal
End Sub

Private Sub Command5_Click()
    DoCmd.Close acForm, "Register"
End Sub

Private Sub Form_Close()
    DoCmd.OpenForm "Menu", acViewForm, "", "", acNormal
End Sub
