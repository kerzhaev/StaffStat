Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =12730
    DatasheetFontHeight =11
    ItemSuffix =18
    Right =18390
    Bottom =12120
    RecSrcDt = Begin
        0x71deebbc227be640
    End
    DatasheetFontName ="Calibri"
    OnActivate ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
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
            TextFontCharSet =204
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            TextFontCharSet =204
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
            BorderLineStyle =0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =4251
            Name ="ОбластьДанных"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    Left =340
                    Top =300
                    Width =8516
                    Height =720
                    TabIndex =3
                    ForeColor =4210752
                    Name ="btnFullSync"
                    Caption ="Full Update"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Run full import, sync, and index update"
                    GridlineColor =10921638

                    LayoutCachedLeft =340
                    LayoutCachedTop =300
                    LayoutCachedWidth =8856
                    LayoutCachedHeight =1020
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =360
                    Top =1140
                    Width =2000
                    Height =240
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lblManualControls"
                    Caption ="Manual Controls"
                    GridlineColor =10921638
                    LayoutCachedLeft =360
                    LayoutCachedTop =1140
                    LayoutCachedWidth =2360
                    LayoutCachedHeight =1380
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =340
                    Top =1440
                    Width =2750
                    Height =576
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnImport"
                    Caption ="Import"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Run manual import"
                    GridlineColor =10921638

                    LayoutCachedLeft =340
                    LayoutCachedTop =1440
                    LayoutCachedWidth =3090
                    LayoutCachedHeight =2016
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3190
                    Top =1440
                    Width =2750
                    Height =576
                    TabIndex =8
                    ForeColor =4210752
                    Name ="cmdHealthCheck"
                    Caption ="Health Check"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Run data integrity check; optionally export errors to Excel"
                    GridlineColor =10921638

                    LayoutCachedLeft =3190
                    LayoutCachedTop =1440
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =2016
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6066
                    Top =1417
                    Width =2000
                    Height =576
                    TabIndex =9
                    ForeColor =4210752
                    Name ="cmdSettings"
                    Caption ="Settings"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open settings (organization, import path, log level)"
                    GridlineColor =10921638

                    LayoutCachedLeft =6066
                    LayoutCachedTop =1417
                    LayoutCachedWidth =8066
                    LayoutCachedHeight =1993
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =340
                    Top =2700
                    Width =2750
                    Height =576
                    TabIndex =5
                    ForeColor =4210752
                    Name ="btnCreateIndexes"
                    Caption ="Create Indexes"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Create performance indexes (one-time operation)"
                    GridlineColor =10921638

                    LayoutCachedLeft =340
                    LayoutCachedTop =2700
                    LayoutCachedWidth =3090
                    LayoutCachedHeight =3276
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =340
                    Top =2070
                    Width =2750
                    Height =576
                    TabIndex =4
                    ForeColor =4210752
                    Name ="btnAnalyze"
                    Caption ="Analyze"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Run manual analysis"
                    GridlineColor =10921638

                    LayoutCachedLeft =340
                    LayoutCachedTop =2070
                    LayoutCachedWidth =3090
                    LayoutCachedHeight =2646
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =3500
                    Top =1140
                    Width =5967
                    Height =240
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lblStatus"
                    Caption ="Log opened."
                    GridlineColor =10921638
                    LayoutCachedLeft =3500
                    LayoutCachedTop =1140
                    LayoutCachedWidth =9467
                    LayoutCachedHeight =1380
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8220
                    Top =1417
                    Width =1985
                    Height =576
                    TabIndex =6
                    ForeColor =4210752
                    Name ="btnOpenLog"
                    Caption ="Open Log"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open change history log"
                    GridlineColor =10921638

                    LayoutCachedLeft =8220
                    LayoutCachedTop =1417
                    LayoutCachedWidth =10205
                    LayoutCachedHeight =1993
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10230
                    Top =1417
                    Width =2500
                    Height =576
                    TabIndex =10
                    ForeColor =4210752
                    Name ="btnFindDuplicates"
                    Caption ="Поиск дубликатов"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Find duplicate personnel records (same FullName and BirthDate)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10230
                    LayoutCachedTop =1417
                    LayoutCachedWidth =12730
                    LayoutCachedHeight =1993
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =93
                    Left =3500
                    Top =2070
                    Width =1200
                    Height =240
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lblStartDate"
                    Caption ="Start Date"
                    GridlineColor =10921638
                    LayoutCachedLeft =3500
                    LayoutCachedTop =2070
                    LayoutCachedWidth =4700
                    LayoutCachedHeight =2310
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =4700
                    Top =2070
                    Width =1550
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtStartDate"
                    ControlTipText ="Enter start date"
                    GridlineColor =10921638

                    LayoutCachedLeft =4700
                    LayoutCachedTop =2070
                    LayoutCachedWidth =6250
                    LayoutCachedHeight =2310
                End
                Begin Label
                    OverlapFlags =93
                    Left =3500
                    Top =2380
                    Width =1200
                    Height =240
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lblEndDate"
                    Caption ="End Date"
                    GridlineColor =10921638
                    LayoutCachedLeft =3500
                    LayoutCachedTop =2380
                    LayoutCachedWidth =4700
                    LayoutCachedHeight =2620
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =4700
                    Top =2380
                    Width =1550
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtEndDate"
                    ControlTipText ="Enter end date"
                    GridlineColor =10921638

                    LayoutCachedLeft =4700
                    LayoutCachedTop =2380
                    LayoutCachedWidth =6250
                    LayoutCachedHeight =2620
                End
                Begin Label
                    OverlapFlags =93
                    Left =6350
                    Top =2070
                    Width =2200
                    Height =240
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lblTotalCount"
                    Caption ="Total: 0"
                    GridlineColor =10921638
                    LayoutCachedLeft =6350
                    LayoutCachedTop =2070
                    LayoutCachedWidth =8550
                    LayoutCachedHeight =2310
                End
                Begin Label
                    OverlapFlags =95
                    Left =6350
                    Top =2310
                    Width =2200
                    Height =240
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lblActiveCount"
                    Caption ="Active: 0"
                    GridlineColor =10921638
                    LayoutCachedLeft =6350
                    LayoutCachedTop =2310
                    LayoutCachedWidth =8550
                    LayoutCachedHeight =2550
                End
                Begin Label
                    OverlapFlags =87
                    Left =6350
                    Top =2550
                    Width =2200
                    Height =240
                    BorderColor =8355711
                    ForeColor =32768
                    Name ="lblErrorCount"
                    Caption ="Errors: 0"
                    GridlineColor =10921638
                    LayoutCachedLeft =6350
                    LayoutCachedTop =2550
                    LayoutCachedWidth =8550
                    LayoutCachedHeight =2790
                    ForeThemeColorIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6292
                    Top =3118
                    Width =2200
                    Height =576
                    TabIndex =11
                    ForeColor =4210752
                    Name ="btnSnapshotReport"
                    Caption ="Штатный срез"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Export current personnel list to Excel"
                    GridlineColor =10921638

                    LayoutCachedLeft =6292
                    LayoutCachedTop =3118
                    LayoutCachedWidth =8492
                    LayoutCachedHeight =3694
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3458
                    Top =3118
                    Width =2750
                    Height =576
                    TabIndex =7
                    ForeColor =4210752
                    Name ="btnGenerateChangeReport"
                    Caption ="Changes Report"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Generate changes report"
                    GridlineColor =10921638

                    LayoutCachedLeft =3458
                    LayoutCachedTop =3118
                    LayoutCachedWidth =6208
                    LayoutCachedHeight =3694
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =360
                    Top =360
                    Width =1005
                    Height =210
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Надпись5"
                    Caption =" "
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =360
                    LayoutCachedTop =360
                    LayoutCachedWidth =1365
                    LayoutCachedHeight =570
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
    End
End
CodeBehindForm
' See "uf_Dashboard.cls"
