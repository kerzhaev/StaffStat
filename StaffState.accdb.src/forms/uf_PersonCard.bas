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
    Width =13039
    DatasheetFontHeight =11
    ItemSuffix =21
    Right =17430
    Bottom =11985
    RecSrcDt = Begin
        0x385c834f3d7be640
    End
    DatasheetFontName ="Calibri"
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
            AddColon = NotDefault
            FELineBreak = NotDefault
            TextFontCharSet =204
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
        Begin ListBox
            TextFontCharSet =204
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            TextFontCharSet =204
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
            Height =10488
            Name ="ОбластьДанных"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =680
                    Top =340
                    Width =9127
                    Height =1125
                    FontSize =14
                    FontWeight =800
                    Name ="txtFullName"

                    LayoutCachedLeft =680
                    LayoutCachedTop =340
                    LayoutCachedWidth =9807
                    LayoutCachedHeight =1465
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =793
                    Top =1870
                    Width =4139
                    Height =855
                    TabIndex =1
                    Name ="txtPersonUID"

                    LayoutCachedLeft =793
                    LayoutCachedTop =1870
                    LayoutCachedWidth =4932
                    LayoutCachedHeight =2725
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =907
                    Top =2948
                    Width =3968
                    Height =1125
                    TabIndex =2
                    Name ="txtRank"

                    LayoutCachedLeft =907
                    LayoutCachedTop =2948
                    LayoutCachedWidth =4875
                    LayoutCachedHeight =4073
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1020
                    Top =4365
                    Width =3685
                    Height =1125
                    TabIndex =3
                    Name ="txtPosition"

                    LayoutCachedLeft =1020
                    LayoutCachedTop =4365
                    LayoutCachedWidth =4705
                    LayoutCachedHeight =5490
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6803
                    Top =4762
                    Width =3459
                    Height =1125
                    TabIndex =4
                    Name ="txtStatus"

                    LayoutCachedLeft =6803
                    LayoutCachedTop =4762
                    LayoutCachedWidth =10262
                    LayoutCachedHeight =5887
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =328
                    Top =6234
                    Width =12131
                    Height =4197
                    TabIndex =5
                    Name ="lstHistory"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="1701;1701;2268;2268"

                    LayoutCachedLeft =328
                    LayoutCachedTop =6234
                    LayoutCachedWidth =12459
                    LayoutCachedHeight =10431
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6803
                    Top =2097
                    Width =2324
                    Height =585
                    TabIndex =6
                    Name ="txtDateFrom"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =6803
                    LayoutCachedTop =2097
                    LayoutCachedWidth =9127
                    LayoutCachedHeight =2682
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9524
                    Top =2154
                    Width =2324
                    Height =585
                    TabIndex =7
                    Name ="txtDateTo"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =9524
                    LayoutCachedTop =2154
                    LayoutCachedWidth =11848
                    LayoutCachedHeight =2739
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12075
                    Top =2211
                    Width =737
                    Height =510
                    TabIndex =8
                    Name ="btnResetDates"
                    Caption ="Х"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =12075
                    LayoutCachedTop =2211
                    LayoutCachedWidth =12812
                    LayoutCachedHeight =2721
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6803
                    Top =3231
                    Width =5102
                    Height =1395
                    TabIndex =9
                    Name ="cboFilterHistory"
                    RowSourceType ="Value List"
                    RowSource ="\"Все\";\"RankName\";\"PosName\";\"WorkStatus\";\"Размер_Сапог\";\"Охват_головы\""
                    OnChange ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =6803
                    LayoutCachedTop =3231
                    LayoutCachedWidth =11905
                    LayoutCachedHeight =4626
                End
            End
        End
    End
End
CodeBehindForm
' See "uf_PersonCard.cls"
