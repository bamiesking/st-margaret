Version =21
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =12756
    DatasheetFontHeight =11
    ItemSuffix =15
    Left =1613
    Top =3120
    DatasheetGridlinesColor =15132391
    Filter ="(([Lookup_Enrolled courses].[Activity Name]=\"Joinery\")) AND ([Lookup_Enrolled "
        "courses].[Activity Name]=\"Joinery\")"
    RecSrcDt = Begin
        0xb2a47aeed163e540
    End
    RecordSource ="Register"
    Caption ="Register"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000d43100006701000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    FitToPage =1
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
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="[Service Users].[Enrolled courses].Value"
        End
        Begin BreakLevel
            ControlSource ="Surname"
        End
        Begin BreakLevel
            ControlSource ="Forename"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1016
            BackColor =15064278
            Name ="ReportHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    Left =57
                    Top =57
                    Width =8760
                    Height =510
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label6"
                    Caption ="Register: "
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =8817
                    LayoutCachedHeight =567
                End
                Begin TextBox
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1587
                    Top =56
                    Width =7257
                    Height =563
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Text13"
                    ControlSource ="=[Forms].[Register].[RegSelBox].[Column](1)"
                    GridlineColor =10921638

                    LayoutCachedLeft =1587
                    LayoutCachedTop =56
                    LayoutCachedWidth =8844
                    LayoutCachedHeight =619
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin CommandButton
                    Left =9015
                    Top =113
                    Width =1088
                    Height =398
                    TabIndex =1
                    ForeColor =4210752
                    Name ="Command9"
                    Caption ="Print"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =9015
                    LayoutCachedTop =113
                    LayoutCachedWidth =10103
                    LayoutCachedHeight =511
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
                    Overlaps =1
                End
            End
        End
        Begin PageHeader
            Height =407
            Name ="PageHeaderSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    TextAlign =1
                    Left =342
                    Top =57
                    Width =2280
                    Height =293
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Service Users.Enrolled courses.Value_Label"
                    Caption ="Activity"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Service_Users_Enrolled_courses_Value_Label"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =342
                    LayoutCachedTop =57
                    LayoutCachedWidth =2622
                    LayoutCachedHeight =350
                End
                Begin Label
                    TextAlign =1
                    Left =2964
                    Top =57
                    Width =4560
                    Height =293
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Surname_Label"
                    Caption ="Surname"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =2964
                    LayoutCachedTop =57
                    LayoutCachedWidth =7524
                    LayoutCachedHeight =350
                End
                Begin Label
                    TextAlign =1
                    Left =7581
                    Top =57
                    Width =3882
                    Height =293
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Forename_Label"
                    Caption ="Forename"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =7581
                    LayoutCachedTop =57
                    LayoutCachedWidth =11463
                    LayoutCachedHeight =350
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =359
            Name ="GroupHeader0"
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =342
                    Width =2280
                    Height =302
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Service Users.Enrolled courses.Value"
                    ControlSource ="=[Forms].[Register].[RegSelBox].[Column](1)"
                    EventProcPrefix ="Service_Users_Enrolled_courses_Value"
                    GridlineColor =10921638

                    LayoutCachedLeft =342
                    LayoutCachedWidth =2622
                    LayoutCachedHeight =302
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =12351
                    Top =56
                    Width =291
                    Height =293
                    TabIndex =1
                    BorderColor =16777215
                    ForeColor =4210752
                    Name ="Text11"
                    GridlineColor =10921638

                    LayoutCachedLeft =12351
                    LayoutCachedTop =56
                    LayoutCachedWidth =12642
                    LayoutCachedHeight =349
                    BorderShade =100.0
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =359
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =2964
                    Width =4560
                    Height =302
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Surname"
                    ControlSource ="Surname"
                    GridlineColor =10921638

                    LayoutCachedLeft =2964
                    LayoutCachedWidth =7524
                    LayoutCachedHeight =302
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =7581
                    Width =3882
                    Height =302
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Forename"
                    ControlSource ="Forename"
                    GridlineColor =10921638

                    LayoutCachedLeft =7581
                    LayoutCachedWidth =11463
                    LayoutCachedHeight =302
                End
            End
        End
        Begin PageFooter
            Height =530
            Name ="PageFooterSection"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =57
                    Top =228
                    Width =5040
                    Height =302
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text7"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    GridlineColor =10921638

                    LayoutCachedLeft =57
                    LayoutCachedTop =228
                    LayoutCachedWidth =5097
                    LayoutCachedHeight =530
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6423
                    Top =228
                    Width =5040
                    Height =302
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text8"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6423
                    LayoutCachedTop =228
                    LayoutCachedWidth =11463
                    LayoutCachedHeight =530
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
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

Private Sub Command9_Click()
    Text11.SetFocus
    Command9.Visible = False
    DoCmd.PrintOut , , , , 1
    Command9.Visible = True
    Command9.SetFocus
End Sub
