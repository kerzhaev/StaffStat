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
    Width =9467
    DatasheetFontHeight =11
    ItemSuffix =15
    Right =25320
    Bottom =12120
    RecSrcDt = Begin
        0x71deebbc227be640
    End
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
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =3070
            Name ="ОбластьДанных"
            AutoHeight =1
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
                    OverlapFlags =87
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
                    OverlapFlags =93
                    Left =340
                    Top =1440
                    Width =2750
                    Height =576
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
                    OverlapFlags =93
                    Left =340
                    Top =2700
                    Width =2750
                    Height =576
                    TabIndex =2
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
                    TabIndex =1
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
                    OverlapFlags =87
                    Left =3500
                    Top =1140
                    Width =5967
                    Height =240
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lblStatus"
                    Caption ="Ready."
                    GridlineColor =10921638
                    LayoutCachedLeft =3500
                    LayoutCachedTop =1140
                    LayoutCachedWidth =9467
                    LayoutCachedHeight =1380
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3500
                    Top =1440
                    Width =2750
                    Height =576
                    TabIndex =3
                    ForeColor =4210752
                    Name ="btnOpenLog"
                    Caption ="Open Log"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open change history log"
                    GridlineColor =10921638

                    LayoutCachedLeft =3500
                    LayoutCachedTop =1440
                    LayoutCachedWidth =6250
                    LayoutCachedHeight =2016
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =87
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
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =4700
                    Top =2070
                    Width =1550
                    Height =240
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
                    OverlapFlags =87
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
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =4700
                    Top =2380
                    Width =1550
                    Height =240
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =3500
                    Top =2700
                    Width =2750
                    Height =576
                    TabIndex =4
                    ForeColor =4210752
                    Name ="btnGenerateChangeReport"
                    Caption ="Changes Report"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Generate changes report"
                    GridlineColor =10921638

                    LayoutCachedLeft =3500
                    LayoutCachedTop =2700
                    LayoutCachedWidth =6250
                    LayoutCachedHeight =3276
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
