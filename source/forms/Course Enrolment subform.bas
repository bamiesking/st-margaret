Version =21
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10209
    DatasheetFontHeight =11
    ItemSuffix =11
    Left =7170
    Top =3435
    Right =20370
    Bottom =7545
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0x8184d82a2d6de540
    End
    RecordSource ="Course Enrolment"
    Caption ="Course Enrolment subform"
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
            Height =0
            BackColor =15064278
            Name ="FormHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
        End
        Begin Section
            Height =3192
            Name ="Detail"
            AutoHeight =255
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =1140
                    Width =1620
                    Height =330
                    ColumnWidth =1860
                    ColumnOrder =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Enrolment Date"
                    ControlSource ="Enrolment Date"
                    EventProcPrefix ="Enrolment_Date"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =1140
                    LayoutCachedWidth =4512
                    LayoutCachedHeight =1470
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =1140
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Enrolment Date_Label"
                            Caption ="Enrolment Date"
                            EventProcPrefix ="Enrolment_Date_Label"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =1140
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =1470
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =1539
                    Width =1620
                    Height =330
                    ColumnWidth =1965
                    ColumnOrder =3
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Completion Date"
                    ControlSource ="Completion Date"
                    EventProcPrefix ="Completion_Date"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =1539
                    LayoutCachedWidth =4512
                    LayoutCachedHeight =1869
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =1539
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Completion Date_Label"
                            Caption ="Completion Date"
                            EventProcPrefix ="Completion_Date_Label"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =1539
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =1869
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =1938
                    Width =7260
                    Height =1140
                    ColumnWidth =7230
                    ColumnOrder =4
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Notes"
                    ControlSource ="Notes"
                    GridlineColor =10921638
                    TextFormat =1

                    LayoutCachedLeft =2892
                    LayoutCachedTop =1938
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =3078
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =1938
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Notes_Label"
                            Caption ="Notes"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =1938
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =2268
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AutoExpand = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =2891
                    Top =737
                    Height =315
                    ColumnWidth =1830
                    ColumnOrder =1
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="Course"
                    ControlSource ="Course"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Activities].[ID], [Activities].[Activity Name] FROM Activities ORDER BY "
                        "[Activity Name]; "
                    ColumnWidths ="0;1440"
                    GridlineColor =10921638

                    LayoutCachedLeft =2891
                    LayoutCachedTop =737
                    LayoutCachedWidth =4592
                    LayoutCachedHeight =1052
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =340
                            Top =737
                            Width =2490
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label10"
                            Caption ="Course"
                            GridlineColor =10921638
                            LayoutCachedLeft =340
                            LayoutCachedTop =737
                            LayoutCachedWidth =2830
                            LayoutCachedHeight =1052
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
