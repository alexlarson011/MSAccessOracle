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
    ItemSuffix =12
    Left =-23988
    Top =3012
    Right =-6768
    Bottom =15984
    RecSrcDt = Begin
        0x9d7a7cfb941fe640
    End
    Caption ="Change Oracle Password"
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
        Begin Section
            Height =4020
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            BackShade =95.0
            Begin
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
                    Width =4380
                    Height =360
                    FontSize =12
                    FontWeight =700
                    Name ="txtAppName"
                    ControlSource ="appDisplayName"

                    LayoutCachedLeft =240
                    LayoutCachedTop =180
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =540
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
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =240
                    Top =540
                    Width =4860
                    Height =600
                    Name ="txtHeader"
                    ControlSource ="headerText"

                    LayoutCachedLeft =240
                    LayoutCachedTop =540
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =1140
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =0
                    ForeTint =100.0
                    GridlineThemeColorIndex =1
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1860
                    Top =1320
                    Width =2460
                    Height =300
                    Name ="txtUserName"

                    LayoutCachedLeft =1860
                    LayoutCachedTop =1320
                    LayoutCachedWidth =4320
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
                            Width =1530
                            Height =300
                            Name ="lblUserName"
                            Caption ="User ID"
                            LayoutCachedLeft =240
                            LayoutCachedTop =1320
                            LayoutCachedWidth =1770
                            LayoutCachedHeight =1620
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1860
                    Top =1680
                    Width =2460
                    Height =300
                    TabIndex =1
                    Name ="txtOldPassword"
                    InputMask ="Password"

                    LayoutCachedLeft =1860
                    LayoutCachedTop =1680
                    LayoutCachedWidth =4320
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
                            Width =1530
                            Height =300
                            Name ="lblOldPassword"
                            Caption ="Current Password"
                            LayoutCachedLeft =240
                            LayoutCachedTop =1680
                            LayoutCachedWidth =1770
                            LayoutCachedHeight =1980
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1860
                    Top =2040
                    Width =2460
                    Height =300
                    TabIndex =2
                    Name ="txtNewPassword"
                    InputMask ="Password"

                    LayoutCachedLeft =1860
                    LayoutCachedTop =2040
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =2340
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =2040
                            Width =1530
                            Height =300
                            Name ="lblNewPassword"
                            Caption ="New Password"
                            LayoutCachedLeft =240
                            LayoutCachedTop =2040
                            LayoutCachedWidth =1770
                            LayoutCachedHeight =2340
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1860
                    Top =2400
                    Width =2460
                    Height =300
                    TabIndex =3
                    Name ="txtConfirmPassword"
                    InputMask ="Password"

                    LayoutCachedLeft =1860
                    LayoutCachedTop =2400
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =2700
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =2400
                            Width =1530
                            Height =300
                            Name ="lblConfirmPassword"
                            Caption ="Confirm Password"
                            LayoutCachedLeft =240
                            LayoutCachedTop =2400
                            LayoutCachedWidth =1770
                            LayoutCachedHeight =2700
                        End
                    End
                End
                Begin Label
                    OverlapFlags =87
                    Left =240
                    Top =2820
                    Width =3300
                    Height =300
                    Name ="lblDSNName"
                    Caption ="DSN"
                    ControlTipText ="Oracle DSN"
                    LayoutCachedLeft =240
                    LayoutCachedTop =2820
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =3120
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =2820
                    Top =3300
                    Width =1980
                    Height =300
                    TabIndex =4
                    Name ="btnUpdatePassword"
                    Caption ="Update Password"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2820
                    LayoutCachedTop =3300
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =3600
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4080
                    Top =3660
                    Width =1200
                    Height =300
                    TabIndex =5
                    Name ="btnCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =4080
                    LayoutCachedTop =3660
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =3960
                    Overlaps =1
                End
            End
        End
    End
End
CodeBehindForm
' See "frmOraclePasswordChange.cls"
