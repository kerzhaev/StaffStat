Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11565
    DatasheetFontHeight =11
    ItemSuffix =29
    Right =18390
    Bottom =12120
    RecSrcDt = Begin
        0xa95bb271d47ce640
    End
    Caption ="Настройки системы"
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
            FontName ="Segoe UI"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =60.0
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
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            TextFontCharSet =204
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Segoe UI"
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
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
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
            FontName ="Segoe UI"
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
            BorderLineStyle =0
        End
        Begin ComboBox
            AddColon = NotDefault
            TextFontCharSet =204
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Segoe UI"
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Tab
            TextFontCharSet =204
            Width =5103
            Height =3402
            FontSize =11
            FontName ="Segoe UI"
            ThemeFontIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =3
            BackThemeColorIndex =1
            BackShade =85.0
            BorderLineStyle =0
            BorderThemeColorIndex =2
            BorderTint =60.0
            HoverThemeColorIndex =1
            PressedThemeColorIndex =1
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin Page
            Width =1701
            Height =1701
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =10920
            Name ="ОбластьДанных"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextFontFamily =0
                    Top =9921
                    Width =2802
                    Height =794
                    ForeColor =4210752
                    Name ="cmdSave"
                    Caption ="Сохранить"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedTop =9921
                    LayoutCachedWidth =2802
                    LayoutCachedHeight =10715
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextFontFamily =0
                    Left =3345
                    Top =9921
                    Width =2802
                    Height =794
                    TabIndex =1
                    ForeColor =4210752
                    Name ="cmdCancel"
                    Caption ="Отмена"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    GridlineColor =10921638

                    LayoutCachedLeft =3345
                    LayoutCachedTop =9921
                    LayoutCachedWidth =6147
                    LayoutCachedHeight =10715
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                End
                Begin Tab
                    OverlapFlags =85
                    TextFontCharSet =177
                    TextFontFamily =0
                    Left =120
                    Top =120
                    Width =7470
                    Height =9705
                    TabIndex =2
                    Name ="tabSettings"
                    FontName ="Calibri Light"
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =7590
                    LayoutCachedHeight =9825
                    BackColor =14277081
                    BorderColor =11573124
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    ForeColor =4210752
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =188
                            Top =593
                            Width =7320
                            Height =9150
                            BorderColor =10921638
                            Name ="pgGeneral"
                            Caption ="Основные настройки"
                            GridlineColor =10921638
                            LayoutCachedLeft =188
                            LayoutCachedTop =593
                            LayoutCachedWidth =7508
                            LayoutCachedHeight =9743
                            Begin
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextFontFamily =0
                                    IMESentenceMode =3
                                    Left =396
                                    Top =963
                                    Width =5896
                                    Height =570
                                    BorderColor =10921638
                                    ForeColor =4210752
                                    Name ="txtOrganizationName"
                                    FontName ="Calibri"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =396
                                    LayoutCachedTop =963
                                    LayoutCachedWidth =6292
                                    LayoutCachedHeight =1533
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextFontFamily =0
                                    IMESentenceMode =3
                                    Left =396
                                    Top =1700
                                    Width =5896
                                    Height =690
                                    BorderColor =10921638
                                    ForeColor =4210752
                                    Name ="txtImportFolderPath"
                                    FontName ="Calibri"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =396
                                    LayoutCachedTop =1700
                                    LayoutCachedWidth =6292
                                    LayoutCachedHeight =2390
                                End
                                Begin CheckBox
                                    OverlapFlags =215
                                    Left =396
                                    Top =2500
                                    Width =239
                                    Height =332
                                    BorderColor =10921638
                                    Name ="chkAutoCheckEnabled"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =396
                                    LayoutCachedTop =2500
                                    LayoutCachedWidth =635
                                    LayoutCachedHeight =2832
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =0
                                            TextFontFamily =0
                                            Left =626
                                            Top =2470
                                            Width =2925
                                            Height =315
                                            BorderColor =8355711
                                            ForeColor =6710886
                                            Name ="Надпись7"
                                            Caption ="Автопроверка после импорта"
                                            FontName ="Calibri"
                                            GridlineColor =10921638
                                            LayoutCachedLeft =626
                                            LayoutCachedTop =2470
                                            LayoutCachedWidth =3551
                                            LayoutCachedHeight =2785
                                        End
                                    End
                                End
                                Begin ComboBox
                                    RowSourceTypeInt =1
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextFontFamily =0
                                    IMESentenceMode =3
                                    Left =396
                                    Top =3514
                                    Width =5102
                                    Height =1665
                                    BorderColor =10921638
                                    ForeColor =3484194
                                    Name ="cboLogLevel"
                                    RowSourceType ="Value List"
                                    RowSource ="DEBUG;INFO;ERROR"
                                    FontName ="Calibri"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =396
                                    LayoutCachedTop =3514
                                    LayoutCachedWidth =5498
                                    LayoutCachedHeight =5179
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =195
                            Top =600
                            Width =7320
                            Height =9150
                            BorderColor =10921638
                            Name ="pgMapping"
                            Caption ="Маппинг импорта"
                            GridlineColor =10921638
                            LayoutCachedLeft =195
                            LayoutCachedTop =600
                            LayoutCachedWidth =7515
                            LayoutCachedHeight =9750
                            Begin
                                Begin ListBox
                                    ColumnHeads = NotDefault
                                    OverlapFlags =247
                                    ColumnCount =3
                                    Left =268
                                    Top =1093
                                    Width =6240
                                    Height =5770
                                    ForeColor =4210752
                                    BorderColor =10921638
                                    Name ="lstMapping"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT MappingID, ExcelHeader, TargetField FROM tbl_Import_Mapping WHERE Profile"
                                        "ID = 1 ORDER BY ExcelHeader; "
                                    ColumnWidths ="0;2501;2501"
                                    FontName ="Calibri"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =268
                                    LayoutCachedTop =1093
                                    LayoutCachedWidth =6508
                                    LayoutCachedHeight =6863
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextFontFamily =0
                                    IMESentenceMode =3
                                    Left =2686
                                    Top =7213
                                    Width =3500
                                    Height =390
                                    BorderColor =10921638
                                    ForeColor =4210752
                                    Name ="txtExcelHeader"
                                    FontName ="Calibri"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =2686
                                    LayoutCachedTop =7213
                                    LayoutCachedWidth =6186
                                    LayoutCachedHeight =7603
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextFontFamily =0
                                            Left =238
                                            Top =7213
                                            Width =2280
                                            Height =315
                                            BorderColor =8355711
                                            ForeColor =6710886
                                            Name ="lblExcel"
                                            Caption ="Заголовок Excel"
                                            FontName ="Calibri"
                                            GridlineColor =10921638
                                            LayoutCachedLeft =238
                                            LayoutCachedTop =7213
                                            LayoutCachedWidth =2518
                                            LayoutCachedHeight =7528
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    RowSourceTypeInt =1
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextFontFamily =0
                                    IMESentenceMode =3
                                    Left =2686
                                    Top =7683
                                    Width =3500
                                    Height =390
                                    BorderColor =10921638
                                    ForeColor =3484194
                                    Name ="cboTargetField"
                                    RowSourceType ="Value List"
                                    RowSource ="\"BankAccountNumber\";\"BankControlKey\";\"BankKey\";\"BirthDate_Text\";\"BootSi"
                                        "ze\";\"CalculatedBy\";\"CalculationUnit\";\"ChildrenCount\";\"Citizenship\";\"Co"
                                        "ntractEndDate\";\"ContractKind\";\"ContractMonths\";\"ContractStartDate\";\"Cont"
                                        "ractType\";\"ContractYears\";\"DismissalDate\";\"EmployeeAge\";\"EmployeeGroup\""
                                        ";\"EmploymentStatus\";\"EventReason\";\"EventType\";\"FullName\";\"Gender\";\"He"
                                        "adSize\";\"MaritalStatus\";\"Nationality\";\"OrderDate_LS\";\"OrderDate_Text\";\""
                                        "OrderNumber\";\"OrderNumber1\";\"Payee\";\"PersonnelDivision\";\"PersonnelDivisi"
                                        "on1\";\"PersonUID\";\"Position\";\"PositionOrderIssuer\";\"SalaryGrade\";\"Sourc"
                                        "eID\";\"StaffPosition\";\"StaffPosition1\";\"ValidFromDate\";\"ValidToDate\";\"V"
                                        "US\";\"Address\""
                                    FontName ="Calibri"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =2686
                                    LayoutCachedTop =7683
                                    LayoutCachedWidth =6186
                                    LayoutCachedHeight =8073
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =0
                                            TextFontFamily =0
                                            Left =238
                                            Top =7683
                                            Width =2280
                                            Height =315
                                            BorderColor =8355711
                                            ForeColor =6710886
                                            Name ="lblDB"
                                            Caption ="Поле в базе"
                                            FontName ="Calibri"
                                            GridlineColor =10921638
                                            LayoutCachedLeft =238
                                            LayoutCachedTop =7683
                                            LayoutCachedWidth =2518
                                            LayoutCachedHeight =7998
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextFontFamily =0
                                    Left =295
                                    Top =8517
                                    Width =2202
                                    Height =794
                                    ForeColor =4210752
                                    Name ="btnAddMapping"
                                    Caption ="Добавить связь"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Calibri"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =295
                                    LayoutCachedTop =8517
                                    LayoutCachedWidth =2497
                                    LayoutCachedHeight =9311
                                    BackColor =14461583
                                    BorderColor =14461583
                                    HoverColor =15189940
                                    PressedColor =9917743
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextFontFamily =0
                                    Left =2595
                                    Top =8517
                                    Width =2202
                                    Height =794
                                    ForeColor =4210752
                                    Name ="btnDeleteMapping"
                                    Caption ="Удалить связь"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Calibri"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =2595
                                    LayoutCachedTop =8517
                                    LayoutCachedWidth =4797
                                    LayoutCachedHeight =9311
                                    BackColor =14461583
                                    BorderColor =14461583
                                    HoverColor =15189940
                                    PressedColor =9917743
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextFontFamily =0
                                    Left =4943
                                    Top =8517
                                    Width =2202
                                    Height =794
                                    ForeColor =4210752
                                    Name ="btnReSeedMapping"
                                    Caption ="Восстановить по умолчанию"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Calibri"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =4943
                                    LayoutCachedTop =8517
                                    LayoutCachedWidth =7145
                                    LayoutCachedHeight =9311
                                    BackColor =14461583
                                    BorderColor =14461583
                                    HoverColor =15189940
                                    PressedColor =9917743
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextFontFamily =0
                                    Left =238
                                    Top =636
                                    Width =2400
                                    Height =315
                                    BorderColor =8355711
                                    ForeColor =6710886
                                    Name ="lblImportMapping"
                                    Caption ="Маппинг импорта"
                                    FontName ="Calibri"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =238
                                    LayoutCachedTop =636
                                    LayoutCachedWidth =2638
                                    LayoutCachedHeight =951
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =195
                            Top =600
                            Width =7320
                            Height =9150
                            BorderColor =10921638
                            Name ="pgMaintenance"
                            Caption ="Обслуживание"
                            GridlineColor =10921638
                            LayoutCachedLeft =195
                            LayoutCachedTop =600
                            LayoutCachedWidth =7515
                            LayoutCachedHeight =9750
                            Begin
                                Begin CommandButton
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextFontFamily =0
                                    Left =238
                                    Top =1544
                                    Width =2802
                                    Height =794
                                    ForeColor =4210752
                                    Name ="cmdCreateBackup"
                                    Caption ="Создать резервную копию"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Calibri"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =238
                                    LayoutCachedTop =1544
                                    LayoutCachedWidth =3040
                                    LayoutCachedHeight =2338
                                    BackColor =14461583
                                    BorderColor =14461583
                                    HoverColor =15189940
                                    PressedColor =9917743
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextFontFamily =0
                                    Left =3583
                                    Top =1544
                                    Width =2802
                                    Height =794
                                    ForeColor =4210752
                                    Name ="cmdClearLogs"
                                    Caption ="Очистить журнал проверки"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Calibri"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =3583
                                    LayoutCachedTop =1544
                                    LayoutCachedWidth =6385
                                    LayoutCachedHeight =2338
                                    BackColor =14461583
                                    BorderColor =14461583
                                    HoverColor =15189940
                                    PressedColor =9917743
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextFontFamily =0
                                    Left =453
                                    Top =2664
                                    Width =2802
                                    Height =794
                                    ForeColor =4210752
                                    Name ="btnRunHealthCheck"
                                    Caption ="Запустить проверку данных"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Calibri"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =453
                                    LayoutCachedTop =2664
                                    LayoutCachedWidth =3255
                                    LayoutCachedHeight =3458
                                    BackColor =14461583
                                    BorderColor =14461583
                                    HoverColor =15189940
                                    PressedColor =9917743
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextFontFamily =0
                                    Left =295
                                    Top =920
                                    Width =4500
                                    Height =345
                                    BorderColor =8355711
                                    ForeColor =6710886
                                    Name ="lblMaintenance"
                                    Caption ="Обслуживание базы данных"
                                    FontName ="Calibri"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =295
                                    LayoutCachedTop =920
                                    LayoutCachedWidth =4795
                                    LayoutCachedHeight =1265
                                End
                            End
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "uf_Settings.cls"
