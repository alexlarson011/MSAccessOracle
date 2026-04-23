Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5580
    DatasheetFontHeight =11
    ItemSuffix =19
    Left =-28140
    Top =3024
    Right =-22308
    Bottom =6432
    RecSrcDt = Begin
        0x9d7a7cfb941fe640
    End
    Caption ="Login"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
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
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
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
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =10921638
            ForeColor =4210752
            FontName ="Calibri"
            AsianLineBreak =1
            GridlineColor =10921638
            BorderShade =65.0
            ThemeFontIndex =1
            ForeTint =75.0
            GridlineShade =65.0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
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
            Height =3420
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            BackShade =95.0
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1260
                    Top =1320
                    Width =2220
                    Height =300
                    Name ="txtUser"

                    LayoutCachedLeft =1260
                    LayoutCachedTop =1320
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =1620
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =1320
                            Width =983
                            Height =300
                            Name ="lblUser"
                            Caption ="User ID"
                            LayoutCachedLeft =240
                            LayoutCachedTop =1320
                            LayoutCachedWidth =1223
                            LayoutCachedHeight =1620
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1260
                    Top =1680
                    Width =3959
                    Height =300
                    TabIndex =1
                    Name ="txtPass"
                    InputMask ="Password"

                    LayoutCachedLeft =1260
                    LayoutCachedTop =1680
                    LayoutCachedWidth =5219
                    LayoutCachedHeight =1980
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =1680
                            Width =983
                            Height =300
                            Name ="lblPass"
                            Caption ="Password"
                            LayoutCachedLeft =240
                            LayoutCachedTop =1680
                            LayoutCachedWidth =1223
                            LayoutCachedHeight =1980
                        End
                    End
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =4020
                    Top =2040
                    Width =1200
                    Height =300
                    TabIndex =2
                    Name ="btnLogin"
                    Caption ="Login"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4020
                    LayoutCachedTop =2040
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =2340
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4020
                    Top =2400
                    Width =1200
                    Height =300
                    TabIndex =3
                    Name ="btnChangePassword"
                    Caption ="Change Password"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4020
                    LayoutCachedTop =2400
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =2700
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4020
                    Top =3000
                    Width =1200
                    Height =300
                    TabIndex =4
                    Name ="btnExit"
                    Caption ="Exit"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4020
                    LayoutCachedTop =3000
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =3300
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =240
                    Top =3000
                    Width =1200
                    Height =300
                    TabIndex =5
                    Name ="btnHelp"
                    Caption ="Help"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =240
                    LayoutCachedTop =3000
                    LayoutCachedWidth =1440
                    LayoutCachedHeight =3300
                    Overlaps =1
                End
                Begin Image
                    PictureType =2
                    Left =5280
                    Top =1680
                    Width =300
                    Height =300
                    Name ="imgHelp"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="helpIconText"
                    Picture ="question25px"

                    LayoutCachedLeft =5280
                    LayoutCachedTop =1680
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =1980
                    TabIndex =9
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =240
                    Top =2640
                    Width =3600
                    Height =300
                    TabIndex =6
                    Name ="txtVersion"
                    ControlSource ="versionText"

                    LayoutCachedLeft =240
                    LayoutCachedTop =2640
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =2940
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =0
                    ForeTint =60.0
                    GridlineThemeColorIndex =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =240
                    Top =540
                    Width =3779
                    Height =720
                    TabIndex =7
                    Name ="txtHeader"
                    ControlSource ="headerText"

                    LayoutCachedLeft =240
                    LayoutCachedTop =540
                    LayoutCachedWidth =4019
                    LayoutCachedHeight =1260
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =0
                    ForeTint =100.0
                    GridlineThemeColorIndex =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =240
                    Top =180
                    Width =3780
                    Height =360
                    FontSize =12
                    FontWeight =700
                    TabIndex =8
                    Name ="txtAppName"
                    ControlSource ="appDisplayName"

                    LayoutCachedLeft =240
                    LayoutCachedTop =180
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =540
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =0
                    ForeTint =100.0
                    GridlineThemeColorIndex =1
                End
                Begin Label
                    OverlapFlags =87
                    Left =240
                    Top =2340
                    Width =1140
                    Height =300
                    Name ="lblDSNName"
                    Caption ="dsnName"
                    OnDblClick ="[Event Procedure]"
                    ControlTipText ="DSN Connection Name"
                    LayoutCachedLeft =240
                    LayoutCachedTop =2340
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =2640
                End
            End
        End
    End
End
CodeBehindForm
' See "frmLogin.cls"
