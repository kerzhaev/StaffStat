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
    ItemSuffix =7
    Right =25320
    Bottom =12120
    RecSrcDt = Begin
        0xeb61691d3d7be640
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
        Begin Section
            Height =7540
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =566
                    Top =737
                    Width =8278
                    Height =1125
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtFilter"
                    OnKeyDown ="[Event Procedure]"
                    OnKeyUp ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =737
                    LayoutCachedWidth =8844
                    LayoutCachedHeight =1862
                End
                Begin Label
                    OverlapFlags =247
                    Top =737
                    Width =675
                    Height =315
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Label1"
                    Caption ="Search:"
                    GridlineColor =10921638
                    LayoutCachedTop =737
                    LayoutCachedWidth =675
                    LayoutCachedHeight =1052
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =680
                    Top =6462
                    Width =2438
                    Height =681
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnSearch"
                    Caption ="Search"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =680
                    LayoutCachedTop =6462
                    LayoutCachedWidth =3118
                    LayoutCachedHeight =7143
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5045
                    Top =6406
                    Width =3062
                    Height =680
                    TabIndex =2
                    ForeColor =4210752
                    Name ="btnClear"
                    Caption ="Clear"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =5045
                    LayoutCachedTop =6406
                    LayoutCachedWidth =8107
                    LayoutCachedHeight =7086
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3168
                    Top =6462
                    Height =681
                    TabIndex =4
                    ForeColor =4210752
                    Name ="btnExportExcel"
                    Caption ="Export to Excel"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3168
                    LayoutCachedTop =6462
                    LayoutCachedWidth =4869
                    LayoutCachedHeight =7143
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =566
                    Top =2211
                    Width =8108
                    Height =3004
                    TabIndex =3
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lstResults"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;2835;1701;2835"
                    OnDblClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =2211
                    LayoutCachedWidth =8674
                    LayoutCachedHeight =5215
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =2211
                            Width =915
                            Height =315
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Label5"
                            Caption ="Results:"
                            GridlineColor =10921638
                            LayoutCachedTop =2211
                            LayoutCachedWidth =915
                            LayoutCachedHeight =2526
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "uf_Search.cls"
