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
    ItemSuffix =8
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
                    Top =510
                    Width =3756
                    Height =576
                    ForeColor =4210752
                    Name ="btnImport"
                    Caption ="Загрузить"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Найти далее"
                    GridlineColor =10921638

                    LayoutCachedLeft =340
                    LayoutCachedTop =510
                    LayoutCachedWidth =4096
                    LayoutCachedHeight =1086
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
                    Left =283
                    Top =2494
                    Width =3756
                    Height =576
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnOpenLog"
                    Caption ="Журнал"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Найти далее"
                    GridlineColor =10921638

                    LayoutCachedLeft =283
                    LayoutCachedTop =2494
                    LayoutCachedWidth =4039
                    LayoutCachedHeight =3070
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
                    Left =5100
                    Top =510
                    Width =3756
                    Height =576
                    TabIndex =3
                    ForeColor =4210752
                    Name ="btnCreateIndexes"
                    Caption ="Создать индексы"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Create performance indexes (one-time operation)"
                    GridlineColor =10921638

                    LayoutCachedLeft =5100
                    LayoutCachedTop =510
                    LayoutCachedWidth =8856
                    LayoutCachedHeight =1086
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
                    Left =226
                    Top =1474
                    Width =3756
                    Height =576
                    TabIndex =2
                    ForeColor =4210752
                    Name ="btnAnalyze"
                    Caption ="Анализ"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Найти далее"
                    GridlineColor =10921638

                    LayoutCachedLeft =226
                    LayoutCachedTop =1474
                    LayoutCachedWidth =3982
                    LayoutCachedHeight =2050
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
                    Left =5612
                    Top =1077
                    Width =3855
                    Height =1757
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lblStatus"
                    Caption ="в"
                    GridlineColor =10921638
                    LayoutCachedLeft =5612
                    LayoutCachedTop =1077
                    LayoutCachedWidth =9467
                    LayoutCachedHeight =2834
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
