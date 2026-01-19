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
    ItemSuffix =14
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
        Begin Section
            Height =8163
            Name ="ОбластьДанных"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =87
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
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =340
                            Width =675
                            Height =315
                            Name ="Надпись1"
                            Caption ="Поле0"
                            LayoutCachedTop =340
                            LayoutCachedWidth =675
                            LayoutCachedHeight =655
                        End
                    End
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
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =1870
                            Width =675
                            Height =315
                            Name ="Надпись5"
                            Caption ="Поле4"
                            LayoutCachedTop =1870
                            LayoutCachedWidth =675
                            LayoutCachedHeight =2185
                        End
                    End
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
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =2948
                            Width =675
                            Height =315
                            Name ="Надпись7"
                            Caption ="Поле6"
                            LayoutCachedTop =2948
                            LayoutCachedWidth =675
                            LayoutCachedHeight =3263
                        End
                    End
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
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =4365
                            Width =675
                            Height =315
                            Name ="Надпись9"
                            Caption ="Поле8"
                            LayoutCachedTop =4365
                            LayoutCachedWidth =675
                            LayoutCachedHeight =4680
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1133
                    Top =5839
                    Width =3459
                    Height =1125
                    TabIndex =4
                    Name ="txtStatus"

                    LayoutCachedLeft =1133
                    LayoutCachedTop =5839
                    LayoutCachedWidth =4592
                    LayoutCachedHeight =6964
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =5839
                            Width =780
                            Height =315
                            Name ="Надпись11"
                            Caption ="Поле10"
                            LayoutCachedTop =5839
                            LayoutCachedWidth =780
                            LayoutCachedHeight =6154
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =6746
                    Top =2154
                    Width =5216
                    Height =4932
                    TabIndex =5
                    Name ="lstHistory"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="1701;1701;2268;2268"

                    LayoutCachedLeft =6746
                    LayoutCachedTop =2154
                    LayoutCachedWidth =11962
                    LayoutCachedHeight =7086
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5045
                            Top =2154
                            Width =1020
                            Height =315
                            Name ="Надпись13"
                            Caption ="Список12:"
                            LayoutCachedLeft =5045
                            LayoutCachedTop =2154
                            LayoutCachedWidth =6065
                            LayoutCachedHeight =2469
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "uf_PersonCard.cls"
