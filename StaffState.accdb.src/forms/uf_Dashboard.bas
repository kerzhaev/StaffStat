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
    Width =9807
    DatasheetFontHeight =11
    ItemSuffix =5
    Right =25575
    Bottom =11985
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
        Begin Section
            Height =5952
            Name ="ОбластьДанных"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =340
                    Top =510
                    Width =3756
                    Height =576
                    Name ="btnImport"
                    Caption ="Кнопка0"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Найти далее"

                    LayoutCachedLeft =340
                    LayoutCachedTop =510
                    LayoutCachedWidth =4096
                    LayoutCachedHeight =1086
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =283
                    Top =2494
                    Width =3756
                    Height =576
                    TabIndex =1
                    Name ="btnOpenLog"
                    Caption ="Кнопка0"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Найти далее"

                    LayoutCachedLeft =283
                    LayoutCachedTop =2494
                    LayoutCachedWidth =4039
                    LayoutCachedHeight =3070
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =226
                    Top =1474
                    Width =3756
                    Height =576
                    TabIndex =2
                    Name ="btnAnalyze"
                    Caption ="Кнопка0"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Найти далее"

                    LayoutCachedLeft =226
                    LayoutCachedTop =1474
                    LayoutCachedWidth =3982
                    LayoutCachedHeight =2050
                End
                Begin Label
                    OverlapFlags =85
                    Left =5612
                    Top =1077
                    Width =3855
                    Height =1757
                    Name ="lblStatus"
                    Caption ="в"
                    LayoutCachedLeft =5612
                    LayoutCachedTop =1077
                    LayoutCachedWidth =9467
                    LayoutCachedHeight =2834
                End
            End
        End
    End
End
CodeBehindForm
' See "uf_Dashboard.cls"
